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

# 앱이 실제로 사용하는 모든 패키지를 한 번에 확인한다.
#  - 엑셀 읽기에 openpyxl이 필요하므로 반드시 포함해야 한다(누락 시 '.xlsx 읽기' 에러).
#  - 친구 검증/입력/이미지는 접근성(AX)·AppKit 기반(AppKit, ApplicationServices).
#  - (구버전의 Quartz/Vision은 더 이상 사용하지 않으므로 설치/확인하지 않는다.)
REQUIRED_IMPORTS="import pandas, openpyxl, pyperclip, flask, AppKit, ApplicationServices"
python3 -c "$REQUIRED_IMPORTS" 2>/dev/null
DEPS_OK=$?

if [ $DEPS_OK -ne 0 ]; then
    echo "📥 패키지 설치 중... (최초 1회만 필요, 수 분 소요)"
    echo ""

    pip3 install --upgrade pip --quiet
    pip3 install pandas openpyxl pyperclip flask pyobjc-framework-Cocoa pyobjc-framework-ApplicationServices --quiet
    if [ $? -ne 0 ]; then
        echo "❌ 패키지 설치에 실패했습니다. 인터넷 연결을 확인하고 다시 실행해주세요."
        read -p "아무 키나 누르면 종료합니다..."
        exit 1
    fi

    # 설치 후 재확인 (부분 설치 방지)
    python3 -c "$REQUIRED_IMPORTS" 2>/dev/null
    if [ $? -ne 0 ]; then
        echo "❌ 일부 패키지가 여전히 누락되었습니다. 다시 실행하거나 아래를 직접 설치해주세요:"
        echo "   pip3 install pandas openpyxl pyperclip flask pyobjc-framework-Cocoa pyobjc-framework-ApplicationServices"
        read -p "아무 키나 누르면 종료합니다..."
        exit 1
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
