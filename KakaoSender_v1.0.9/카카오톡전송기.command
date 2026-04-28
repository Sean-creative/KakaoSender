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

# 가상환경 디렉토리
VENV_DIR=".venv"

# 가상환경이 없으면 생성
if [ ! -d "$VENV_DIR" ]; then
    echo "🔧 가상환경 생성 중... (최초 1회만 필요)"
    python3 -m venv "$VENV_DIR"
    if [ $? -ne 0 ]; then
        echo "❌ 가상환경 생성에 실패했습니다."
        read -p "아무 키나 누르면 종료합니다..."
        exit 1
    fi
    echo "✅ 가상환경 생성 완료!"
    echo ""
fi

# 가상환경 활성화
source "$VENV_DIR/bin/activate"

# 필요한 패키지 확인 및 설치
echo "📦 필요한 패키지 확인 중..."

# 기본 패키지 확인
python3 -c "import pandas, pyperclip, flask" 2>/dev/null
BASIC_OK=$?

# OCR 패키지 확인 (친구 검증용)
python3 -c "import Quartz, Vision" 2>/dev/null
OCR_OK=$?

if [ $BASIC_OK -ne 0 ] || [ $OCR_OK -ne 0 ]; then
    echo "📥 패키지 설치 중... (최초 1회만 필요, 수 분 소요)"
    echo ""
    
    # pip 업그레이드
    pip3 install --upgrade pip --quiet
    
    if [ $BASIC_OK -ne 0 ]; then
        echo "   - 기본 패키지 설치 중..."
        pip3 install pandas pyperclip openpyxl flask --quiet
        if [ $? -ne 0 ]; then
            echo "❌ 기본 패키지 설치에 실패했습니다."
            read -p "아무 키나 누르면 종료합니다..."
            exit 1
        fi
    fi
    
    if [ $OCR_OK -ne 0 ]; then
        echo "   - OCR 패키지 설치 중... (시간이 다소 걸립니다)"
        pip3 install pyobjc-framework-Quartz pyobjc-framework-Vision --quiet
        if [ $? -ne 0 ]; then
            echo "❌ OCR 패키지 설치에 실패했습니다."
            read -p "아무 키나 누르면 종료합니다..."
            exit 1
        fi
    fi
    
    echo ""
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
