#!/bin/bash
set -e
cd "$(dirname "$0")"

echo "========================================================"
echo "  PPT-style 个人 PPT 风格工坊 · 一键安装程序 (macOS)"
echo "========================================================"
echo "  提示：如双击遇到“无法验证开发者”，请右键本文件 → 打开。"
echo ""

PY_CMD=""
for cand in "/opt/homebrew/bin/python3" "/usr/local/bin/python3" "$(command -v python3 2>/dev/null || true)" "/usr/bin/python3"; do
  if [ -n "$cand" ] && [ -x "$cand" ]; then
    # 真运行验证，防止误中未装 Xcode Command Line Tools 的 /usr/bin/python3 假桩弹窗
    if "$cand" -c "import sys; assert sys.version_info[0]==3" >/dev/null 2>&1; then
      PY_CMD="$cand"
      break
    fi
  fi
done

if [ -z "$PY_CMD" ]; then
  echo "[提示] 未检测到真实可用的 Python 3 运行环境。"
  echo "你可以任选一种方式准备环境后再双击本文件："
  echo "  1. 从 https://www.python.org/downloads/ 下载官方 macOS 安装包安装；"
  echo "  2. 或打开「终端」执行：xcode-select --install"
  read -p "按回车键关闭窗口..."
  exit 1
fi

echo "[1/2] 正在使用 Python 环境：$PY_CMD"
"$PY_CMD" -X utf8 "./install.py"

echo ""
echo "[2/2] 安装已完成！请【新开一个对话窗口】（或重启你的 AI 软件），先发这句做 3 秒验活："
echo "    检查 PPT-style 是否就绪"
echo ""
echo "看到【PPT-style ... 已激活就绪】回复后，再发：帮我建 PPT 风格工坊"
echo "如果新开对话后反馈找不到技能，请让当前 AI 检查它实际读取的全局技能目录并同步。"
echo ""
read -p "按回车键关闭窗口..."
