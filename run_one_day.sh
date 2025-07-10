#!/bin/bash

# 사용법 안내
usage() {
  echo "사용법: $0 [시작일] [종료일]"
  echo "예시:  $0 2025-01-01 2025-03-31"
  echo "       (인자 없을 경우 기본값: 시작일=2025-01-01, 종료일=오늘)"
}

# 인자 처리
start_date=${1:-"2025-01-01"}        # 인자가 없으면 기본 시작일
end_date=${2:-$(date +%F)}           # 인자가 없으면 오늘

# 유효성 검사
if ! date -d "$start_date" >/dev/null 2>&1; then
  echo "❌ 시작일 형식 오류: $start_date"
  usage; exit 1
fi

if ! date -d "$end_date" >/dev/null 2>&1; then
  echo "❌ 종료일 형식 오류: $end_date"
  usage; exit 1
fi

echo "실행 범위: $start_date ~ $end_date"

current_date="$start_date"

while [[ "$current_date" < "$end_date" || "$current_date" == "$end_date" ]]; do
  echo "▶ 실행 중: ./fund_mail_one_day --date $current_date"
  ./fund_mail_one_day --date "$current_date"

  # 10초 대기 (필요시 조정)
  sleep 3

  # 다음 날짜로 증가
  current_date=$(date -I -d "$current_date + 1 day")
done

echo "✅ 모든 작업 완료."
