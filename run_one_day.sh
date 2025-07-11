#!/bin/bash

# 사용법 안내
usage() {
  echo "사용법: $0 <시작일> <종료일>"
  echo "예시:  $0 2025-01-01 2025-03-31"
  echo "주의: 시작일과 종료일 모두 필수이며, 시작일은 종료일보다 과거여야 합니다."
}

# 인자 개수 검증
if [ $# -ne 2 ]; then
  echo "❌ 인자 개수 오류: 정확히 2개의 인자가 필요합니다."
  usage; exit 1
fi

# 인자 처리
start_date=$1
end_date=$2

# 유효성 검사
if ! date -d "$start_date" >/dev/null 2>&1; then
  echo "❌ 시작일 형식 오류: $start_date"
  usage; exit 1
fi

if ! date -d "$end_date" >/dev/null 2>&1; then
  echo "❌ 종료일 형식 오류: $end_date"
  usage; exit 1
fi

# 날짜 순서 검증 (시작일이 종료일보다 과거여야 함)
if [[ "$start_date" > "$end_date" ]]; then
  echo "❌ 날짜 순서 오류: 시작일($start_date)이 종료일($end_date)보다 늦습니다."
  usage; exit 1
fi

echo "실행 범위: $start_date ~ $end_date"

# 사용자 확인
echo -n "진행하시겠습니까? (Ctrl+C로 취소, 다른 키로 계속): "
read -r confirmation

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
