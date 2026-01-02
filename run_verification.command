#!/bin/bash
# 카카오톡 친구 검증기 (Verification)

# 스크립트가 있는 디렉토리로 이동
cd "$(dirname "$0")"

echo "======================================"
echo "    카카오톡 친구 검증기 (OCR)"
echo "======================================"
echo ""

# Python 확인
if ! command -v python3 &> /dev/null; then
    echo "❌ Python3가 설치되어 있지 않습니다."
    exit 1
fi

# 필요한 패키지 (pyobjc, pyperclip, pandas, openpyxl) 확인
if ! python3 -c "import Quartz, Vision, pyperclip, pandas, openpyxl" 2>/dev/null; then
    echo "📦 필요한 패키지를 설치합니다 (pyobjc, pyperclip, pandas, openpyxl)..."
    python3 -m pip install pyobjc pyperclip pandas openpyxl --quiet
fi

# Python 스크립트 실행 (엑셀 파일 처리)
echo "🚀 엑셀 파일 기반 자동화 검증을 시작합니다..."
python3 verify_friend_search.py
EXIT_CODE=$?

echo ""
if [ $EXIT_CODE -eq 0 ]; then
    echo "✅ 작업 완료."
else
    echo "❌ 작업 중 오류 발생."
fi

# 터미널이 바로 꺼지지 않게 대기 (더블 클릭 실행 시 유용)
echo ""
# read -p "종료하려면 아무 키나 누르세요..." -n 1 -s
