#!/bin/bash
# linux용임
# 플랫폼 확인
OS="$(uname -s)"
EXE_NAME="fund_mail_one_day"
ENTRY_SCRIPT="src/main_one_day.py"

# OS에 따라 Python 명령 설정 및 출력 파일명 결정
if [[ "$OS" == "Linux"* || "$OS" == "Darwin"* ]]; then
    PYTHON_CMD="python3"
    FINAL_EXE_NAME="$EXE_NAME"
else
    PYTHON_CMD="python"
    FINAL_EXE_NAME="$EXE_NAME.exe"
fi

# Python 설치 확인
if ! command -v $PYTHON_CMD &> /dev/null; then
    echo "❌ $PYTHON_CMD 이(가) 설치되어 있지 않습니다."
    exit 1
fi

# PyInstaller 설치 확인
if ! $PYTHON_CMD -m PyInstaller --version &> /dev/null; then
    echo "📦 pyinstaller가 설치되어 있지 않습니다. 설치 중..."
    $PYTHON_CMD -m pip install pyinstaller
fi

# 빌드 실행
echo "🛠️  [$OS] 빌드 중..."
$PYTHON_CMD -m PyInstaller --name "$EXE_NAME" --onefile "$ENTRY_SCRIPT"

# 결과 안내
if [[ -f "dist/$FINAL_EXE_NAME" ]]; then
    echo "✅ 빌드 완료: dist/$FINAL_EXE_NAME"
else
    echo "❌ 빌드 실패: dist/$FINAL_EXE_NAME 가 생성되지 않았습니다."
fi
