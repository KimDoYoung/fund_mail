#!/usr/bin/env bash
set -euo pipefail

# ----- 설정 -----
ENTRYPOINT="src/main_once.py"          # PyInstaller 진입점
APP_NAME="fund_mail_once"              # exe 이름
DIST_DIR="dist"                   # 산출물 폴더
VENV_DIR=".venv"                  # (선택) 가상환경 사용 시
# -----------------

echo "▶️  빌드 시작 : $ENTRYPOINT → $DIST_DIR/$APP_NAME.exe"

# 1) 깨끗한 dist 폴더 확보
rm -rf "$DIST_DIR"
mkdir -p "$DIST_DIR"

# 2) (선택) 가상환경 안에서 PyInstaller 실행
if [[ -d "$VENV_DIR" ]]; then
  source "$VENV_DIR/Scripts/activate"
fi

# 3) 단일-파일 exe 생성
pyinstaller \
  --onefile \
  --noconfirm \
  --name "$APP_NAME" \
  --distpath "$DIST_DIR" \
  "$ENTRYPOINT"

# 4) .env, 추가 리소스 복사
cp .env "$DIST_DIR"/

echo "✅  빌드 완료 : $DIST_DIR/$APP_NAME.exe"
