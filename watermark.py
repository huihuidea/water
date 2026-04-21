import os
import tkinter as tk
from tkinter import filedialog, messagebox
from threading import Thread

from document_watermark import process_document


class DocumentToolFrame(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg="#fbf7ef", padx=16, pady=16)
        self.files = []
        self.save_folder = ""
        self._build()

    def _build(self):
        hero = tk.Frame(self, bg="#f5e7c8", padx=14, pady=12)
        hero.pack(fill="x", pady=(0, 12))
        tk.Label(hero, text="Word / PDF 文档去水印", font=("Arial", 15, "bold"), bg="#f5e7c8").pack(anchor="w")
        tk.Label(
            hero,
            text="处理 .docx 和 .pdf。适合页眉页脚水印，以及一部分正文下层浮动文字/图片水印。",
            wraplength=620,
            bg="#f5e7c8",
            anchor="w",
            justify="left",
        ).pack(fill="x", pady=(4, 0))

        body = tk.Frame(self, bg="#fbf7ef")
        body.pack(fill="both", expand=True)

        left = tk.LabelFrame(body, text="文档处理", bg="#fbf7ef", padx=12, pady=10)
        left.pack(side="left", fill="both", expand=True, padx=(0, 8))
        tk.Button(left, text="选择 Word / PDF 文档", command=self.choose).pack(fill="x", pady=4)
        tk.Button(left, text="选择文档输出目录", command=self.choose_save).pack(fill="x", pady=4)
        tk.Button(left, text="开始处理文档水印", command=self.start, bg="#9a6700", fg="white").pack(fill="x", pady=(10, 0))

        self.files_var = tk.StringVar(value="未选择文档")
        self.save_var = tk.StringVar(value="未选择保存目录")
        self.note_var = tk.StringVar(
            value="PDF 目前主要支持 annotation 型水印；直接画进页面内容流的复杂底层水印暂不保证全部命中。"
        )

        right = tk.LabelFrame(body, text="支持说明", bg="#fbf7ef", padx=12, pady=10)
        right.pack(side="left", fill="both", expand=True, padx=(8, 0))
        tk.Label(right, textvariable=self.files_var, fg="#444", bg="#fbf7ef", anchor="w", justify="left").pack(fill="x", pady=2)
        tk.Label(right, textvariable=self.save_var, fg="#444", bg="#fbf7ef", anchor="w", justify="left").pack(fill="x", pady=2)
        tk.Label(
            right,
            textvariable=self.note_var,
            fg="#666",
            bg="#fbf7ef",
            wraplength=560,
            anchor="w",
            justify="left",
        ).pack(fill="x", pady=(10, 0))
        tk.Label(
            right,
            text="输出文件名会自动加上 no_wm_ 前缀，原文件不会被覆盖。",
            fg="#666",
            bg="#fbf7ef",
            wraplength=560,
            anchor="w",
            justify="left",
        ).pack(fill="x", pady=(10, 0))

    def choose(self):
        fs = filedialog.askopenfilenames(
            title="选择 Word / PDF 文档",
            filetypes=[("文档文件", "*.docx *.pdf")],
        )
        if fs:
            self.files = list(fs)
            self.files_var.set(f"已选择 {len(self.files)} 个文档")

    def choose_save(self):
        d = filedialog.askdirectory(title="选择保存文件夹")
        if d:
            self.save_folder = d
            self.save_var.set(f"保存目录: {d}")

    def start(self):
        if not self.files:
            messagebox.showwarning("注意", "先选择文档")
            return
        if not self.save_folder:
            messagebox.showwarning("注意", "先选择保存文件夹")
            return
        Thread(target=self._process_documents, daemon=True).start()

    def _process_documents(self):
        ok = 0
        failed = []
        for src in self.files:
            name = os.path.basename(src)
            save_path = os.path.join(self.save_folder, f"no_wm_{name}")
            try:
                process_document(src, dst=save_path)
                ok += 1
            except Exception as e:
                failed.append(f"{name} -> {e}")

        def notify():
            if failed:
                names = "\n".join(failed[:10])
                extra = "\n..." if len(failed) > 10 else ""
                messagebox.showwarning(
                    "完成(部分失败)",
                    f"共{len(self.files)}个文档，成功{ok}个，失败{len(failed)}个\n\n失败示例:\n{names}{extra}",
                )
            else:
                messagebox.showinfo("完成", f"共{len(self.files)}个文档，成功{ok}个")

        self.after(0, notify)


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Word / PDF 去水印工具")
        self.root.geometry("760x480")
        self.root.minsize(700, 440)

        top = tk.Frame(root, bg="#ffffff", padx=14, pady=14)
        top.pack(fill="x")
        tk.Label(top, text="Word / PDF 去水印工具", font=("Arial", 18, "bold"), bg="#ffffff").pack(anchor="w")
        tk.Label(
            top,
            text="当前桌面版仅保留文档去水印功能。",
            fg="#555",
            bg="#ffffff",
        ).pack(anchor="w", pady=(4, 0))

        content = tk.Frame(root, bg="#e9eef5")
        content.pack(fill="both", expand=True)
        DocumentToolFrame(content).pack(fill="both", expand=True)

        footer = tk.Frame(root, bg="#ffffff", padx=14, pady=10)
        footer.pack(fill="x")
        tk.Label(
            footer,
            text="支持 .docx 与 .pdf。桌面版无需浏览器，可直接打包为 macOS / Windows 应用。",
            fg="#666",
            bg="#ffffff",
            anchor="w",
            justify="left",
        ).pack(fill="x")


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
