#!/usr/bin/env bash
set -euo pipefail

# ----- 설정 -----
DIST_DIR="dist"
SRC_DIR="src"
UV_FILES=(pyproject.toml uv.lock requirements.txt README.md)   # uv sync 용
# -----------------

echo "▶️  서비스용 폴더 생성 : $DIST_DIR"

# 1) 깨끗한 dist 확보
rm -rf "$DIST_DIR"
mkdir -p "$DIST_DIR/$SRC_DIR"

# 2) 파이썬 소스 전체 복사 (rsync 대신 cp -a)
#    - -a : 보존(archive) + 재귀
cp -a "$SRC_DIR"/. "$DIST_DIR/$SRC_DIR/"

# 3) __pycache__ 제거 (선택)
find "$DIST_DIR" -type d -name "__pycache__" -exec rm -rf {} +

# 4) .env 복사
cp .env "$DIST_DIR"/
cp manage_fund_mail.bat "$DIST_DIR"/
# 5) uv sync 관련 파일 복사
for f in "${UV_FILES[@]}"; do
  [[ -f "$f" ]] && cp "$f" "$DIST_DIR"/
done

echo "ℹ️  Windows 서비스 등록은 'nssm' 또는 'sc.exe' 를 사용하세요."
echo "   예) nssm install fund_mail \"C:\\Python\\python.exe\" \"-m\" \"$SRC_DIR.main\""
echo "✅  서비스용 디렉터리 준비 완료 : $DIST_DIR"
