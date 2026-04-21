#!/bin/zsh
set -e

cd "/Users/hejinhui/Desktop/python"

if ! command -v python3 >/dev/null 2>&1; then
  osascript -e 'display dialog "未找到 python3，请先安装 Python 3" buttons {"确定"} default button "确定"'
  exit 1
fi

python3 "watermark.py"
