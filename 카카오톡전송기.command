#!/bin/bash
# 카카오톡 자동 전송기 실행 스크립트 (웹 버전)

# 스크립트가 있는 디렉토리로 이동
cd "$(dirname "$0")"

echo "======================================"
echo "  카카오톡 자동 전송기 (웹 버전)"
echo "======================================"
echo ""

# Python 확인
if ! command -v python3 &> /dev/null; then
    echo "❌ Python3가 설치되어 있지 않습니다."
    echo "   Xcode Command Line Tools를 설치해주세요:"
    echo "   xcode-select --install"
    read -p "아무 키나 누르면 종료합니다..."
    exit 1
fi

# 필요한 패키지 확인 및 설치
echo "📦 필요한 패키지 확인 중..."
python3 -c "import pandas, pyperclip, flask" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "📥 패키지 설치 중... (최초 1회만 필요)"
    pip3 install pandas pyperclip openpyxl flask --quiet
    echo "✅ 패키지 설치 완료!"
fi

echo ""
echo "🌐 웹 브라우저에서 프로그램이 열립니다..."
echo "   종료하려면 이 창을 닫거나 Ctrl+C를 누르세요."
echo ""

# 웹 버전 실행
python3 kakao_web.py

echo ""
echo "프로그램이 종료되었습니다."
