#!/bin/zsh
set -e

cd "/Users/hejinhui/Desktop/python"

APP_PATH="dist/AIWatermarkTool.app"
DMG_PATH="dist/AIWatermarkTool-mac.dmg"
VOL_NAME="AIWatermarkTool"

if [ ! -d "$APP_PATH" ]; then
  echo "未找到 $APP_PATH，请先执行 ./build_desktop.command"
  exit 1
fi

rm -f "$DMG_PATH"

hdiutil create -volname "$VOL_NAME" -srcfolder "$APP_PATH" -ov -format UDZO "$DMG_PATH"

osascript -e 'display dialog "DMG 打包完成，输出文件在 dist/AIWatermarkTool-mac.dmg" buttons {"确定"} default button "确定"'
