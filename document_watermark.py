"""
Remove common *document* watermarks (Word .docx / PDF).

- DOCX: strips typical header/footer watermark XML and also removes likely
  body-level floating watermark shapes / drawings embedded in document.xml.
- PDF: removes PDF annotation subtype Watermark when present (PyMuPDF).

Embedded PDF watermarks drawn as normal page content (not annotations)
cannot be removed reliably without rewriting the page content stream.
"""

from __future__ import annotations

import io
import os
import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

_NS_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_NS_MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent"
_NS_V = "{urn:schemas-microsoft-com:vml}"
_NS_WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
_NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_NS_PIC = "{http://schemas.openxmlformats.org/drawingml/2006/picture}"


def _qn_w(local: str) -> str:
    return _NS_W + local


def _xml_lower(elem: ET.Element) -> str:
    try:
        return ET.tostring(elem, encoding="unicode", default_namespace=None).lower()
    except Exception:
        return ET.tostring(elem, encoding="utf-8", errors="ignore").decode("utf-8", errors="ignore").lower()


def _is_watermarkish_fragment(s: str) -> bool:
    s = s.lower()
    # Explicit producer markers.
    markers = (
        "wordpicturewatermark",
        "wordtextwatermark",
        "textwatermark",
        "picturewatermark",
    )
    return any(m in s for m in markers)


def _has_any(text: str, needles: tuple[str, ...]) -> bool:
    return any(n in text for n in needles)


def _style_value(style: str, key: str) -> str | None:
    prefix = key.lower() + ":"
    for part in style.split(";"):
        item = part.strip()
        if item.lower().startswith(prefix):
            return item[len(prefix):].strip().lower()
    return None


def _looks_like_body_watermark_fragment(s: str) -> bool:
    s = s.lower()
    if _is_watermarkish_fragment(s):
        return True

    score = 0

    # Typical Word floating watermark clues.
    if _has_any(s, ("behinddoc=\"1\"", "wrapnone", "mso-position-horizontal:center", "mso-position-vertical:center")):
        score += 2
    if _has_any(s, ("rotation:", "mso-rotate", "style=\"rotation:", "style='rotation:")):
        score += 1
    if _has_any(s, ("opacity:", "fillopacity", "alpha", "transparency")):
        score += 1
    if _has_any(s, ("z-index:-", "z-index: -", "position:absolute", "position: absolute")):
        score += 1
    if _has_any(s, ("fillcolor=", "stroked=\"f\"", "allowincell=\"f\"")):
        score += 1

    # Very large decorative text/picture containers are often watermarks.
    if _has_any(s, ("font-size:4", "font-size:5", "font-size:6", "font-size:7", "font-size:8", "font-size:9")):
        score += 1
    if _has_any(s, ("width:4", "width:5", "width:6", "width:7", "width:8", "width:9")) and _has_any(
        s, ("height:2", "height:3", "height:4", "height:5", "height:6")
    ):
        score += 1

    return score >= 3


def _anchor_behind_doc(anchor: ET.Element) -> bool:
    return anchor.attrib.get("behindDoc") == "1"


def _shape_text(shape: ET.Element) -> str:
    parts: list[str] = []
    for node in shape.iter():
        if node.text and node.text.strip():
            parts.append(node.text.strip())
    return " ".join(parts).lower()


def _looks_like_vml_watermark(shape: ET.Element) -> bool:
    xml = _xml_lower(shape)
    if _looks_like_body_watermark_fragment(xml):
        return True

    style = shape.attrib.get("style", "")
    style_l = style.lower()
    score = 0
    if _has_any(style_l, ("position:absolute", "position: absolute", "rotation:", "z-index:-", "z-index: -")):
        score += 1
    if _has_any(style_l, ("width:4", "width:5", "width:6", "width:7", "width:8", "width:9")):
        score += 1
    if _has_any(style_l, ("height:2", "height:3", "height:4", "height:5", "height:6")):
        score += 1

    opacity = _style_value(style, "opacity")
    if opacity is not None:
        score += 1
        if opacity.endswith("%"):
            try:
                if float(opacity[:-1]) <= 35:
                    score += 1
            except ValueError:
                pass
        elif opacity.replace(".", "", 1).isdigit():
            try:
                if float(opacity) <= 0.35:
                    score += 1
            except ValueError:
                pass

    text = _shape_text(shape)
    if len(text) >= 4 and len(text) <= 80:
        score += 1

    return score >= 4


def _extent_is_large(elem: ET.Element) -> bool:
    cx = elem.attrib.get("cx")
    cy = elem.attrib.get("cy")
    try:
        return int(cx or "0") >= 3_000_000 and int(cy or "0") >= 1_500_000
    except ValueError:
        return False


def _looks_like_drawing_watermark(anchor: ET.Element) -> bool:
    xml = _xml_lower(anchor)
    if _looks_like_body_watermark_fragment(xml):
        return True

    score = 0
    if _anchor_behind_doc(anchor):
        score += 2

    for extent in anchor.iter(_NS_WP + "extent"):
        if _extent_is_large(extent):
            score += 1
            break

    for xfrm in anchor.iter(_NS_A + "xfrm"):
        rot = xfrm.attrib.get("rot")
        try:
            if rot and abs(int(rot)) >= 1_500_000:
                score += 1
                break
        except ValueError:
            pass

    text = ""
    for t in anchor.iter(_NS_A + "t"):
        if t.text and t.text.strip():
            text += t.text.strip() + " "
    text = text.strip().lower()
    if 4 <= len(text) <= 80:
        score += 1

    if any(True for _ in anchor.iter(_NS_PIC + "pic")):
        score += 1

    return score >= 3


def _parent_map(root: ET.Element) -> dict[ET.Element, ET.Element | None]:
    return {child: parent for parent in root.iter() for child in list(parent)}


def _remove_watermark_nodes_from_tree(root: ET.Element) -> int:
    """Remove watermark-related nodes; return approximate removal count."""
    removed = 0

    # Remove in passes so parent pointers stay valid.
    while True:
        parents = _parent_map(root)
        target: ET.Element | None = None
        for ac in root.iter(_NS_MC):
            if _is_watermarkish_fragment(_xml_lower(ac)):
                target = ac
                break
        if target is None:
            for pict in root.iter(_qn_w("pict")):
                if _is_watermarkish_fragment(_xml_lower(pict)):
                    target = pict
                    break
        if target is None:
            for shape in root.iter(_NS_V + "shape"):
                if _looks_like_vml_watermark(shape):
                    target = shape
                    break
        if target is None:
            for anchor in root.iter(_NS_WP + "anchor"):
                if _looks_like_drawing_watermark(anchor):
                    target = anchor
                    break
        if target is None:
            break
        par = parents.get(target)
        if par is None:
            break
        par.remove(target)
        removed += 1

    # Prune empty runs / paragraphs left behind (safe for typical watermark-only blocks).
    parents = _parent_map(root)
    for run in list(root.iter(_qn_w("r"))):
        if len(run) == 0:
            par = parents.get(run)
            if par is not None:
                par.remove(run)
                removed += 1
        parents = _parent_map(root)

    for para in list(root.iter(_qn_w("p"))):
        if len(para) == 0:
            par = parents.get(para)
            if par is not None:
                par.remove(para)
                removed += 1
        parents = _parent_map(root)

    return removed


_DOC_PARTS = (
    "word/document.xml",
    "word/endnotes.xml",
    "word/footnotes.xml",
)


def _iter_header_footer_paths(names: list[str]) -> list[str]:
    out: list[str] = []
    for n in names:
        if (n.startswith("word/header") or n.startswith("word/footer")) and n.endswith(".xml"):
            out.append(n)
    return sorted(set(out))


def remove_docx_watermarks(src: str | Path, dst: str | Path | None = None) -> tuple[Path, str]:
    """
    Copy docx to dst (or temp file) and strip known watermark XML fragments.
    Returns (output_path, log_message).
    """
    src = Path(src)
    if not src.is_file():
        raise FileNotFoundError(str(src))
    if src.suffix.lower() != ".docx":
        raise ValueError("Only .docx is supported (not .doc binary).")

    if dst is None:
        fd, tmp = tempfile.mkstemp(suffix=".docx", prefix="no_wm_")
        os.close(fd)
        dst = Path(tmp)
    else:
        dst = Path(dst)

    shutil.copy2(src, dst)

    total_removed = 0
    touched: list[str] = []

    with zipfile.ZipFile(dst, "r") as zin:
        names = zin.namelist()
        targets = [t for t in list(_DOC_PARTS) + _iter_header_footer_paths(names) if t in names]

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for zinfo in zin.infolist():
                data = zin.read(zinfo.filename)
                if zinfo.filename in targets:
                    try:
                        root = ET.fromstring(data)
                    except ET.ParseError:
                        zout.writestr(zinfo, data)
                        continue
                    n = _remove_watermark_nodes_from_tree(root)
                    if n:
                        total_removed += n
                        touched.append(zinfo.filename)
                    # Standard Office part declaration; Word accepts this on round-trip.
                    out_xml = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + ET.tostring(
                        root, encoding="utf-8", default_namespace=None
                    )
                    new_info = zipfile.ZipInfo(filename=zinfo.filename, date_time=zinfo.date_time)
                    new_info.compress_type = zipfile.ZIP_DEFLATED
                    zout.writestr(new_info, out_xml)
                else:
                    zout.writestr(zinfo, data)

    buf.seek(0)
    with open(dst, "wb") as f:
        f.write(buf.getvalue())

    msg = f"DOCX processed. XML watermark-related nodes removed (approx): {total_removed}."
    if touched:
        msg += " Parts touched: " + ", ".join(touched[:12])
        if len(touched) > 12:
            msg += ", ..."
    else:
        msg += " No known header/footer or body-level watermark blocks were found (file may use a different style)."
    return dst, msg


def remove_pdf_watermarks(src: str | Path, dst: str | Path | None = None) -> tuple[Path, str]:
    """
    Remove PDF Watermark annotations if the producer used them.
    """
    try:
        import fitz  # PyMuPDF
    except ImportError as e:
        raise ImportError("PDF support needs PyMuPDF: pip install pymupdf") from e

    src = Path(src)
    if not src.is_file():
        raise FileNotFoundError(str(src))
    if src.suffix.lower() != ".pdf":
        raise ValueError("Only .pdf is supported.")

    if dst is None:
        fd, tmp = tempfile.mkstemp(suffix=".pdf", prefix="no_wm_")
        os.close(fd)
        dst = Path(tmp)
    else:
        dst = Path(dst)

    doc = fitz.open(src)
    removed = 0
    try:
        wm_type = getattr(fitz, "PDF_ANNOT_WATERMARK", 22)
        for page in doc:
            annot = page.first_annot
            while annot:
                nxt = annot.next
                try:
                    if annot.type[0] == wm_type:
                        page.delete_annot(annot)
                        removed += 1
                except Exception:
                    pass
                annot = nxt
        doc.save(dst, garbage=4, deflate=True, clean=True)
    finally:
        doc.close()

    msg = f"PDF saved. Watermark annotations removed: {removed}."
    if removed == 0:
        msg += " If the watermark is painted as normal page content (not an annotation), it cannot be removed this way."
    return dst, msg


def process_document(path: str | Path, dst: str | Path | None = None) -> tuple[Path, str]:
    """Dispatch by extension."""
    path = Path(path)
    suf = path.suffix.lower()
    if suf == ".docx":
        return remove_docx_watermarks(path, dst=dst)
    if suf == ".pdf":
        return remove_pdf_watermarks(path, dst=dst)
    raise ValueError(f"Unsupported type: {suf}. Use .docx or .pdf.")
