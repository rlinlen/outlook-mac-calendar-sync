#!/bin/bash

# Outlook Calendar to Google Calendar 同步腳本
# 使用方法: 
#   ./sync_outlook_to_google.sh           # 使用預設14天
#   ./sync_outlook_to_google.sh 7         # 同步7天
#   ./sync_outlook_to_google.sh 30        # 同步30天

# 設定預設天數
DAYS=${1:-14}

echo "🚀 開始 Outlook Calendar 到 Google Calendar 同步..."
echo "=================================================="
echo "📅 同步範圍: ${DAYS} 天"

# 檢查是否安裝了 uv
if ! command -v uv &> /dev/null; then
    echo "❌ 錯誤: 找不到 uv 命令"
    echo "請先安裝 uv: curl -LsSf https://astral.sh/uv/install.sh | sh"
    exit 1
fi

# 激活 uv 環境
echo "🔧 激活 uv 環境..."
if [ -f "./.venv/bin/activate" ]; then
    source "./.venv/bin/activate"
    echo "✅ 已載入 venv 環境"
fi

# 確保 uv 可用
if command -v uv &> /dev/null; then
    echo "✅ uv 命令可用"
else
    echo "❌ 激活後仍找不到 uv 命令"
    exit 1
fi

# 步驟 1: 生成 Outlook CSV 檔案
echo "📊 步驟 1: 讀取 Outlook 行事曆資料..."
if uv run ./script/dump_outlook_calendar.py --days ${DAYS}; then
    echo "✅ Outlook 資料讀取成功"
else
    echo "❌ Outlook 資料讀取失敗"
    exit 1
fi

# 步驟 2: 同步到 Google Calendar
echo ""
echo "🔄 步驟 2: 同步到 Google Calendar..."
if uv run ./script/sync_csv_with_google_calendar.py --days ${DAYS}; then
    echo "✅ Google Calendar 同步成功"
else
    echo "❌ Google Calendar 同步失敗"
    exit 1
fi

echo ""
echo "🎉 同步完成！"
echo "📅 請檢查你的 Google Calendar 查看同步結果"
