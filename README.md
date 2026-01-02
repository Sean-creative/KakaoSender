# 📬 카카오톡 자동 전송기 (macOS)

엑셀 파일에서 대상자를 필터링하여 카카오톡 메시지를 자동으로 전송하는 macOS 전용 프로그램입니다.

![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)
![macOS](https://img.shields.io/badge/macOS-only-black.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)

## ✨ 특징

- **웹 기반 인터페이스**: 브라우저에서 간편하게 조작
- **AppleScript 기반 자동화**: 이미지 인식 없이 안정적으로 동작
- **엑셀 필터링**: 조건에 맞는 회원만 자동 추출
- **실시간 로그**: 전송 진행 상황 실시간 확인

## 🎯 필터링 조건

| 조건 | 값 |
|------|---|
| 등록형태 | 이월, 재등록, 신규 |
| 연령 | 20대, 30대 |

## 📋 요구사항

- macOS 10.15 이상
- Python 3.9 이상
- 카카오톡 macOS 앱

## 🚀 설치 및 실행

### 방법 1: 배포 파일 사용 (비개발자 추천)

1. `배포용` 폴더를 다운로드합니다.
2. `카카오톡전송기.command` 파일을 더블클릭합니다.
3. 최초 실행 시 필요한 패키지가 자동으로 설치됩니다.

> ⚠️ "확인되지 않은 개발자" 경고가 뜨면:
> - 파일을 우클릭(Control+클릭) → "열기" 선택
> - 또는 시스템 설정 > 개인정보 보호 및 보안 > "그래도 열기" 클릭

### 방법 2: 직접 실행 (개발자)

```bash
# 저장소 클론
git clone https://github.com/Sean-creative/KakaoSender.git
cd KakaoSender

# 의존성 설치
pip install pandas pyperclip openpyxl flask

# 실행
python kakao_web.py
```

## 📁 엑셀 파일 형식

엑셀 파일에는 다음 컬럼이 필요합니다:

| 컬럼명 | 설명 |
|--------|------|
| 이름 | 회원 이름 (카카오톡 친구 이름과 동일해야 함) |
| 등록형태 | 이월 / 재등록 / 신규 |
| 연령 | 20대 / 30대 / 40대 등 |

## 🔧 메시지 커스터마이징

`kakao_web.py`에서 메시지 템플릿을 수정할 수 있습니다:

```python
MESSAGE_TEMPLATE = "{name}님!\n요청하신 리포트입니다.\n감사합니다."
```

## ⚙️ 동작 원리

1. 엑셀 파일 업로드 및 필터링
2. AppleScript로 카카오톡 활성화
3. 친구 검색 (Cmd+F → 이름 붙여넣기)
4. 채팅방 진입 (↓↓ → Enter)
5. 메시지 전송 (Cmd+V → Enter)
6. 다음 대상자로 이동

## ⚠️ 주의사항

- **전송 중 마우스/키보드 조작 금지**
- 카카오톡이 로그인되어 있어야 합니다
- 화면이 잠기지 않도록 해주세요
- 친구 이름이 카카오톡에 등록된 이름과 정확히 일치해야 합니다

## 📂 프로젝트 구조

```
KakaoSender/
├── README.md
├── kakao_web.py              # 메인 프로그램 (웹 버전)
├── 카카오톡전송기.command      # 실행 스크립트
└── 사용방법.txt               # 사용 설명서
```

## 🙏 참고

- [autokakao](https://github.com/yoonhero/autokakao) - AppleScript 기반 카카오톡 자동화

## 📄 License

MIT License
