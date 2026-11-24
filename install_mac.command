#!/bin/bash
set -e
cd "$(dirname "$0")"

echo "========================================================"
echo "  PPT-style 个人 PPT 风格工坊 · 一键安装程序 (macOS)"
echo "========================================================"
echo ""

PY_CMD=""
for cand in "/opt/homebrew/bin/python3" "/usr/local/bin/python3" "/usr/bin/python3" "$(command -v python3 2>/dev/null || true)"; do
  if [ -n "$cand" ] && [ -x "$cand" ]; then
    PY_CMD="$cand"
    break
  fi
done

if [ -z "$PY_CMD" ]; then
  echo "[提示] 未检测到可用的 Python 3 运行环境。"
  echo "请打开终端执行 xcode-select --install 或从 python.org 安装 Python 3 后再双击本文件。"
  read -p "按回车键关闭窗口..."
  exit 1
fi

"$PY_CMD" -X utf8 "./install.py"

echo ""
echo "安装已完成！请【新开一个对话窗口】（或重启你的 AI 软件），先发这句做 3 秒验活："
echo "    检查 PPT-style 是否就绪"
echo ""
echo "看到【PPT-style ... 已激活就绪】回复后，再发：帮我建 PPT 风格工坊"
echo ""
read -p "按回车键关闭窗口..."
