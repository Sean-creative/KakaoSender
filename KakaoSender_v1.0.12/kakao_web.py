"""
카카오톡 자동 메시지 전송기 (macOS) - 웹 버전
- 브라우저 기반 인터페이스 (tkinter 사용 안함)
- Flask 웹 서버 사용
"""

import os
import sys
import re
import tempfile
import subprocess
import threading
import time
import random
import webbrowser
from queue import Queue, Empty
from datetime import datetime
from typing import Optional, List

import pandas as pd
import pyperclip
from flask import Flask, render_template_string, request, jsonify, Response
import Quartz
import Vision

# 접근성(AX) API — 친구 이름을 OCR 없이 정확한 문자열로 읽기 위한 경로.
# 사용 불가 환경(프레임워크 누락 등)에서는 자동으로 OCR 폴백만 사용한다.
try:
    from AppKit import NSWorkspace
    from ApplicationServices import (
        AXUIElementCreateApplication,
        AXUIElementCopyAttributeValue,
        AXIsProcessTrusted,
    )
    AX_AVAILABLE = True
except Exception:
    AX_AVAILABLE = False

# ============================================================
# 설정
# ============================================================
VERSION = "1.0.12"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# 카카오톡 자동화 튜닝 (다른 맥에서 검색/채팅 진입 안정화)
KAKAOTALK_WINDOW_X = 50
KAKAOTALK_WINDOW_Y = 50
KAKAOTALK_WINDOW_WIDTH = 1400
KAKAOTALK_WINDOW_HEIGHT = 900
SEARCH_RESULT_DOWN_ARROW_COUNT = 2
DECORATED_NAME_MIN_CHARS = 3
SUBSTRING_MATCH_MIN_CHARS = 4
MAX_SEARCH_OCR_ATTEMPTS = 2
SEARCH_INPUT_VERIFY_ATTEMPTS = 2
# 친구 검증 시 접근성(AX) API를 우선 사용하고, 실패하면 기존 OCR로 폴백한다.
USE_AX_VERIFICATION = True
# 카카오톡 번들 식별자 (AX 앱 핸들 탐색용)
KAKAO_BUNDLE_IDS = ('com.kakao.KakaoTalkMac', 'com.kakao.KakaoTalk')
# 검색 결과에서 친구 이름이 담기는 AXStaticText 의 identifier
AX_DISPLAY_NAME_ID = 'Display Name'
OCR_DEBUG_SAVE = True
OCR_DEBUG_DIR = os.path.join(tempfile.gettempdir(), 'kakao_sender_ocr_debug')
CHAT_DELAY_BEFORE_ENTER = 0.55
CHAT_DELAY_AFTER_ENTER = 1.35
# 업로드 파일은 스크립트 폴더가 아닌 시스템 임시 디렉터리에 저장 (폴더명 공백·복사본 경로 등으로 인한 ENOENT 방지)
UPLOAD_TEMP_XLSX = os.path.join(tempfile.gettempdir(), f'kakao_sender_upload_{os.getpid()}.xlsx')

# 선택 가능한 필터 옵션
AVAILABLE_REGISTER_TYPES = ['이월', '재등록', '신규', '이탈', '이탈(단)']
AVAILABLE_AGE_GROUPS = ['10대', '20대', '30대', '40대', '50대', '60대 이상']

# 기본 선택값
DEFAULT_REGISTER_TYPES = ['이월', '재등록', '신규']
DEFAULT_AGE_GROUPS = ['20대', '30대']

# 기본 메시지 템플릿
DEFAULT_MESSAGE_TEMPLATE = "{name}님!\n요청하신 리포트입니다.\n감사합니다."

# Flask 앱
app = Flask(__name__)
# 로그 스트림(SSE) 구독자 목록. 연결마다 독립 큐를 갖고, 로그는 모든 구독자에게
# 동일 메시지를 브로드캐스트한다. 단일 공유 큐를 여러 연결이 나눠 먹어 초기 로그가
# 유실되던 문제를 원천 차단한다.
log_subscribers = []
log_subscribers_lock = threading.Lock()


def _broadcast(item_json: str):
    """이미 JSON 문자열인 이벤트를 모든 SSE 구독자 큐에 전달한다."""
    with log_subscribers_lock:
        subscribers = list(log_subscribers)
    for q in subscribers:
        try:
            q.put(item_json)
        except Exception:
            pass


class _BroadcastQueue:
    """기존 log_queue.put(json_str) 호출부를 그대로 두기 위한 브로드캐스트 어댑터."""
    def put(self, item_json: str):
        _broadcast(item_json)


# 기존 코드의 log_queue.put(...) 호출은 그대로 두고, 동작만 브로드캐스트로 바꾼다.
log_queue = _BroadcastQueue()
is_running = False
stop_requested = False
pause_requested = False
pause_event = threading.Event()
pause_event.set()  # starts unpaused
current_file_path = None
current_register_types = None
current_age_groups = None
current_message_template = None


@app.errorhandler(Exception)
def handle_exception(e):
    """전역 오류 핸들러"""
    return jsonify({'success': False, 'error': str(e)}), 500

# ============================================================
# AppleScript 헬퍼
# ============================================================
def run_applescript(script: str) -> tuple:
    """AppleScript 실행"""
    proc = subprocess.Popen(
        ['osascript', '-'],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE
    )
    out, err = proc.communicate(input=script.encode('utf-8'))
    return proc.returncode, out.decode('utf-8'), err.decode('utf-8')


# 주의: U+24C2~U+1F251 같은 넓은 범위는 한글 음절(U+AC00~)까지 포함해 이름이 전부 지워짐
EMOJI_PATTERN = re.compile(
    "["
    "\U0001F600-\U0001F64F"  # Emoticons
    "\U0001F300-\U0001F5FF"  # Symbols & Pictographs
    "\U0001F680-\U0001F6FF"  # Transport & Map
    "\U0001F1E0-\U0001F1FF"  # Flags
    "\U00002702-\U000027B0"
    "\U0001F900-\U0001F9FF"  # Supplemental Symbols
    "\U0001FA00-\U0001FA6F"  # Chess Symbols
    "\U0001FA70-\U0001FAFF"  # Symbols Extended-A
    "\U00002600-\U000026FF"  # Misc symbols
    "\U0000FE00-\U0000FE0F"  # Variation Selectors
    "\U0000200D"             # Zero Width Joiner
    "\U00002B50"             # Star
    "\U000024C2"             # Ⓜ (단일 기호, 범위 금지)
    "\U0000203C-\U00003299"  # 일부 기호·한자 호환 (한글 음절 U+AC00 미포함)
    "]+", flags=re.UNICODE
)


def normalize_name(name: str) -> str:
    """이름 정규화: 연속 공백을 하나로 + 앞뒤 공백 제거 (동명이인 비교용)"""
    return re.sub(r'\s+', ' ', str(name)).strip()


def normalize_name_for_ocr_match(name: str) -> str:
    """OCR 비교용 이름 정규화: 이모티콘/장식 기호 제거 + 공백 정리."""
    text = EMOJI_PATTERN.sub('', str(name))
    return normalize_name(text)


def comparable_name_length(name: str) -> int:
    """공백을 제외한 비교용 이름 길이."""
    return len(re.sub(r'\s+', '', name))


def name_contains_emoji_or_symbol(name: str) -> bool:
    """이름에 이모티콘/감지용 특수기호가 포함되는지 (전송 차단 판별용)"""
    return bool(EMOJI_PATTERN.search(str(name)))


# ============================================================
# AppleScript 명령어
# ============================================================
SCRIPT_ACTIVATE = '''
tell application "KakaoTalk" to activate
'''

# 검색창 초기화 (다음 검색을 위해)
SCRIPT_RESET_SEARCH = '''
tell application "KakaoTalk" to activate
delay 0.3
tell application "System Events"
    tell process "KakaoTalk"
        set frontmost to true
    end tell
    delay 0.2
    
    -- 1. Esc 3회 (채팅창/검색창/알림 등 모든 레이어 닫기)
    key code 53
    delay 0.3
    key code 53
    delay 0.3
    key code 53
    delay 0.3
    
    -- 2. 친구 목록으로 이동 (Cmd+1)
    keystroke "1" using command down
    delay 0.5
end tell
'''


# ============================================================
# OCR 헬퍼 함수
# ============================================================
def get_kakaotalk_window_id():
    """카카오톡 창 ID 가져오기"""
    options = Quartz.kCGWindowListOptionOnScreenOnly
    window_list = Quartz.CGWindowListCopyWindowInfo(options, Quartz.kCGNullWindowID)
    
    candidates = []
    for window in window_list:
        owner_name = window.get('kCGWindowOwnerName', '')
        window_id = window.get('kCGWindowNumber', 0)
        bounds = window.get('kCGWindowBounds', {})
        
        # 카카오톡 & 어느정도 크기가 있는 메인창
        if 'KakaoTalk' in owner_name or '카카오톡' in owner_name:
            if bounds.get('Width', 0) > 200 and bounds.get('Height', 0) > 200:
                candidates.append(window_id)
                
    return candidates[0] if candidates else None


def wait_for_kakaotalk_window(timeout: float = 3.0, interval: float = 0.3) -> Optional[int]:
    """UI 전환 중 잠깐 사라지는 카카오톡 창이 다시 잡힐 때까지 대기."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        window_id = get_kakaotalk_window_id()
        if window_id:
            return window_id
        time.sleep(interval)
    return get_kakaotalk_window_id()


def capture_and_read(window_id: int) -> List[str]:
    """창 캡처 후 OCR로 텍스트 읽기"""
    # 캡처
    cg_image = Quartz.CGWindowListCreateImage(
        Quartz.CGRectNull,
        Quartz.kCGWindowListOptionIncludingWindow,
        window_id,
        Quartz.kCGWindowImageBoundsIgnoreFraming | Quartz.kCGWindowImageNominalResolution
    )
    
    if not cg_image:
        return []

    # OCR
    request_handler = Vision.VNImageRequestHandler.alloc().initWithCGImage_options_(cg_image, None)
    request = Vision.VNRecognizeTextRequest.alloc().init()
    request.setRecognitionLevel_(Vision.VNRequestTextRecognitionLevelAccurate)
    request.setRecognitionLanguages_(['ko-KR', 'en-US'])
    
    success, error = request_handler.performRequests_error_([request], None)
    
    results_text = []
    if success:
        observations = request.results()
        if observations:
            for obs in observations:
                candidate = obs.topCandidates_(1)[0]
                results_text.append(candidate.string())
                
    return results_text


def save_window_screenshot(window_id: int, label: str) -> Optional[str]:
    """디버그용으로 카카오톡 창 캡처를 PNG로 저장. 실패하면 None 반환."""
    if not OCR_DEBUG_SAVE:
        return None
    try:
        os.makedirs(OCR_DEBUG_DIR, exist_ok=True)
        safe_label = re.sub(r'[^\w가-힣\-]+', '_', str(label), flags=re.UNICODE)[:40] or 'ocr'
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        out_path = os.path.join(OCR_DEBUG_DIR, f'{ts}_{safe_label}.png')
        subprocess.run(
            ['screencapture', '-l', str(window_id), '-x', out_path],
            check=False, timeout=5,
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        if os.path.exists(out_path) and os.path.getsize(out_path) > 0:
            return out_path
    except Exception:
        return None
    return None


def ensure_kakaotalk_ready() -> Optional[int]:
    """카카오톡이 활성화되어 있고 창이 열려있는지 확인, 필요시 재시도"""
    max_retries = 3
    
    for attempt in range(max_retries):
        # 카카오톡 활성화
        script = '''
        tell application "KakaoTalk"
            activate
            delay 0.5
        end tell
        tell application "System Events"
            tell process "KakaoTalk"
                set frontmost to true
                -- 창이 없으면 새 창 열기 시도
                if (count of windows) is 0 then
                    keystroke "n" using command down
                    delay 0.5
                end if
            end tell
        end tell
        '''
        run_applescript(script)
        time.sleep(0.5)
        
        # 창 확인: 카카오톡 UI 전환 중 일시적으로 창 목록이 비는 경우가 있어 짧게 대기
        window_id = wait_for_kakaotalk_window(timeout=2.0)
        if window_id:
            return window_id
        
        time.sleep(1)
    
    return None


def _paste_into_search(name: str, use_keystroke: bool) -> None:
    """검색창 열고 검색어를 붙여넣기. use_keystroke=True면 Cmd+V, False면 메뉴 클릭 사용."""
    pyperclip.copy(name)
    if use_keystroke:
        paste_block = '''
        keystroke "v" using command down
        '''
    else:
        paste_block = '''
        tell process "KakaoTalk"
            set frontmost to true
            try
                click menu item "붙여넣기" of menu "편집" of menu bar 1
            on error
                try
                    click menu item "Paste" of menu "편집" of menu bar 1
                on error
                    try
                        click menu item "Paste" of menu "Edit" of menu bar 1
                    end try
                end try
            end try
        end tell
        '''
    script = f'''
    tell application "KakaoTalk" to activate
    delay 0.3

    tell application "System Events"
        tell process "KakaoTalk"
            set frontmost to true
        end tell
        delay 0.3

        -- 채팅창/검색창 닫기
        key code 53
        delay 0.3

        -- 친구 목록으로 이동 (Cmd+1)
        keystroke "1" using command down
        delay 0.5

        -- 검색창 열기 (Cmd+F)
        key code 3 using command down
        delay 0.5

        -- 기존 검색어 전체 선택 + 삭제
        key code 0 using command down
        delay 0.2
        key code 51
        delay 0.3

        -- 붙여넣기 (메뉴 클릭 또는 Cmd+V)
        {paste_block}
        delay 0.9
    end tell
    '''
    run_applescript(script)


def _read_search_field_text() -> str:
    """검색창에 들어 있는 텍스트를 읽어옴.

    1순위: 접근성(AX) API로 AXSearchField 값을 직접 읽음 (부수효과 없음, 정확).
    2순위: 실패 시 기존 방식(클립보드 경유 Cmd+A → Cmd+C)으로 폴백.
    """
    if AX_AVAILABLE and USE_AX_VERIFICATION:
        try:
            if AXIsProcessTrusted():
                window = _ax_get_main_window(_ax_get_kakao_app_element())
                if window is not None:
                    ax_value = _ax_read_search_field(window)
                    if ax_value is not None:
                        return ax_value
        except Exception:
            pass

    script = '''
    tell application "System Events"
        tell process "KakaoTalk"
            set frontmost to true
            key code 0 using command down
            delay 0.15
            key code 8 using command down
            delay 0.25
        end tell
    end tell
    '''
    run_applescript(script)
    time.sleep(0.2)
    try:
        return pyperclip.paste() or ""
    except Exception:
        return ""


def _move_focus_to_search_results() -> None:
    """검색창에서 아래 화살표로 결과 리스트로 포커스 이동."""
    n_down = SEARCH_RESULT_DOWN_ARROW_COUNT
    script = f'''
    tell application "System Events"
        tell process "KakaoTalk"
            set frontmost to true
        end tell
        repeat with _k from 1 to {n_down}
            key code 125
            delay 0.2
        end repeat
    end tell
    '''
    run_applescript(script)


def search_friend(name: str) -> bool:
    """친구 검색. 검색창에 검색어가 실제로 들어갔는지 검증하고, 실패 시 Cmd+V로 재시도.

    Returns:
        True  - 검색창 입력 검증까지 성공
        False - 두 번 시도해도 검증 실패 (그래도 결과 리스트로 포커스는 이동시켜 둠)
    """
    expected = normalize_name(name)
    verified = False
    for attempt in range(SEARCH_INPUT_VERIFY_ATTEMPTS):
        use_keystroke = attempt > 0  # 1회차는 메뉴 클릭, 재시도부터 Cmd+V
        _paste_into_search(name, use_keystroke=use_keystroke)
        time.sleep(0.3)
        actual = normalize_name(_read_search_field_text())
        if actual == expected:
            verified = True
            break
        method = "Cmd+V" if use_keystroke else "메뉴 클릭"
        log(f"   -> ⚠️ 검색창 입력 확인 실패 ({method} 시도, 실제: '{actual[:30]}'). 재시도합니다.")

    _move_focus_to_search_results()
    return verified


def _looks_like_search_field(candidate: str, query_normalized: str, decorated_query: str) -> bool:
    """OCR 후보가 카카오톡 검색창 텍스트인지 판별.

    검색창 좌측 돋보기 아이콘이 OCR에서 `Q`, `O`, `0` 등 1~2글자로 잘못 읽혀
    `<짧은 prefix> <검색어>` 형태로 잡힌다. 친구 영역 매칭에서 이를 제외해야
    검색창 텍스트가 친구 일치로 오인되지 않는다.
    """
    if not candidate:
        return False
    parts = candidate.strip().split(None, 1)
    if len(parts) != 2:
        return False
    prefix, rest = parts
    if len(prefix) > 2:
        return False
    rest_plain = normalize_name(rest)
    rest_decorated = normalize_name_for_ocr_match(rest)
    return rest_plain == query_normalized or rest_decorated == decorated_query


# ============================================================
# 접근성(AX) 기반 친구 검증
#   - 카카오톡 검색 결과의 친구 이름을 OCR 없이 정확한 문자열로 읽는다.
#   - 검색창(AXSearchField)과 친구행(AXStaticText id='Display Name')이
#     role/identifier로 명확히 구분되어 OCR의 오인식·검색창 혼동이 사라진다.
# ============================================================
def _ax_copy(element, attr):
    """AX 속성 한 개를 안전하게 읽어 반환. 실패 시 None."""
    try:
        err, value = AXUIElementCopyAttributeValue(element, attr, None)
        if err != 0:
            return None
        return value
    except Exception:
        return None


def _ax_get_kakao_app_element():
    """실행 중인 카카오톡의 AX 앱 요소를 반환. 없으면 None."""
    try:
        workspace = NSWorkspace.sharedWorkspace()
        for app in workspace.runningApplications():
            bundle_id = app.bundleIdentifier() or ""
            name = app.localizedName() or ""
            if bundle_id in KAKAO_BUNDLE_IDS or 'KakaoTalk' in name or '카카오톡' in name:
                return AXUIElementCreateApplication(app.processIdentifier())
    except Exception:
        return None
    return None


def _ax_get_main_window(app_element):
    """카카오톡 메인 창 AX 요소를 반환. AXMainWindow → AXWindows[0] → AXFocusedWindow 순."""
    if app_element is None:
        return None
    win = _ax_copy(app_element, "AXMainWindow")
    if win is not None:
        return win
    windows = _ax_copy(app_element, "AXWindows") or []
    if windows:
        return windows[0]
    return _ax_copy(app_element, "AXFocusedWindow")


def _ax_walk(element, max_depth, depth=0):
    """AX 트리를 깊이 우선으로 순회하며 모든 요소를 산출."""
    yield element
    if depth >= max_depth:
        return
    children = _ax_copy(element, "AXChildren") or []
    for child in children:
        yield from _ax_walk(child, max_depth, depth + 1)


def _ax_collect_result_names(window, max_depth: int = 30) -> List[str]:
    """검색 결과 행의 친구 이름(AXStaticText id='Display Name')을 모두 수집."""
    names = []
    for el in _ax_walk(window, max_depth):
        if _ax_copy(el, "AXIdentifier") == AX_DISPLAY_NAME_ID:
            value = _ax_copy(el, "AXValue")
            if value:
                text = normalize_name(str(value))
                if text:
                    names.append(text)
    return names


def _ax_read_search_field(window, max_depth: int = 30) -> Optional[str]:
    """검색창(AXSearchField)에 입력된 텍스트를 읽어 반환. 없으면 None."""
    for el in _ax_walk(window, max_depth):
        if _ax_copy(el, "AXSubrole") == "AXSearchField":
            value = _ax_copy(el, "AXValue")
            if value is not None:
                return str(value)
    return None


def verify_friend_by_ax(name: str) -> bool:
    """접근성(AX) API로 친구 검증. 확인되면 True, 아니면 False.

    OCR과 달리 검색창과 친구행이 role/identifier로 구분되므로
    검색창 텍스트를 친구로 오인할 위험이 없다.
    실패(미확인/읽기 불가) 시 호출부에서 OCR로 폴백한다.
    """
    if not (AX_AVAILABLE and USE_AX_VERIFICATION):
        return False

    if not AXIsProcessTrusted():
        log("   -> ⚠️ 접근성 권한이 없어 AX 검증을 건너뜁니다. (OCR로 진행)")
        return False

    # AX 경로에서 어떤 예외가 나더라도 OCR 폴백을 막지 않도록 전체를 감싼다.
    try:
        app_element = _ax_get_kakao_app_element()
        window = _ax_get_main_window(app_element)
        if window is None:
            return False

        normalized = normalize_name(name)
        decorated_normalized = normalize_name_for_ocr_match(name)

        # 오발송 방지 가드: 검색창에 실제로 이 검색어가 들어가 있을 때만 결과를 신뢰한다.
        # (검색 필터가 안 된 채 전체 목록이 보이는 상태에서 우연히 일치해 잘못 보내는 것을 차단)
        search_value = _ax_read_search_field(window)
        if search_value is None:
            log("   -> ⚠️ AX 검색창을 찾지 못해 AX 검증을 보류합니다. (OCR로 진행)")
            return False
        search_normalized = normalize_name(search_value)
        if search_normalized != normalized and normalize_name_for_ocr_match(search_value) != decorated_normalized:
            log(
                f"   -> ⚠️ AX 검색창 값('{search_normalized[:30]}')이 검색어와 달라 "
                f"AX 검증을 보류합니다. (OCR로 진행)"
            )
            return False

        names = _ax_collect_result_names(window)
        if not names:
            return False

        # 1) 정확 일치
        if normalized in names:
            log(f"   -> ✅ AX 확인됨: '{normalized}'")
            return True

        # 2) 장식기호 제거 후 일치
        decorated_matches = {n for n in names if normalize_name_for_ocr_match(n) == decorated_normalized}
        if len(decorated_matches) == 1:
            log(f"   -> ✅ AX(장식기호 제거) 확인됨: '{next(iter(decorated_matches))}'")
            return True
        if len(decorated_matches) > 1:
            log(f"   -> ⚠️ AX 후보가 여러 개라 오발송 방지를 위해 보류: {', '.join(decorated_matches)}")
            return False

        # 3) 긴 이름에 한해 단일 부분 일치 허용
        if comparable_name_length(decorated_normalized) >= SUBSTRING_MATCH_MIN_CHARS:
            substring_matches = [
                n for n in names
                if normalize_name_for_ocr_match(n) != decorated_normalized
                and decorated_normalized in normalize_name_for_ocr_match(n)
            ]
            if len(substring_matches) == 1:
                log(f"   -> ✅ AX(부분 일치) 확인됨: '{substring_matches[0]}' ⊇ '{decorated_normalized}'")
                return True

        return False
    except Exception as exc:
        log(f"   -> ⚠️ AX 검증 중 예외 발생, OCR로 진행합니다: {exc}")
        return False


def verify_friend_by_ocr(name: str, window_id: int) -> bool:
    """OCR로 친구 검증.

    0) 검색창 텍스트('Q <검색어>' 형태)는 매칭 대상에서 제외 (오발송 방지)
    1) 정확 일치 우선
    2) 실패 시 이모티콘/장식 기호 제거 후 엄격 비교
    3) 그래도 실패 시 긴 이름에 한해 조건부 부분 일치 (정확히 한 후보 줄에 포함될 때만)
    4) 짧은 이름이나 중복 후보는 오발송 방지를 위해 실패 처리
    5) 최종 실패 시 OCR 후보와 캡처 이미지를 디버그용으로 남김
    """
    texts = capture_and_read(window_id)
    normalized = normalize_name(name)
    decorated_normalized = normalize_name_for_ocr_match(name)

    # 0) 검색창 텍스트는 매칭 대상에서 제외 — 검색창은 입력한 검색어가 그대로 보이므로
    #    필터링하지 않으면 '친구가 없는 경우'에도 검색창만 보고 잘못 매칭될 수 있음.
    friend_texts = [
        t for t in texts
        if not _looks_like_search_field(t, normalized, decorated_normalized)
    ]

    # 1) 정확 일치
    exact_matches = set()
    for t in friend_texts:
        if normalized == normalize_name(t):
            exact_matches.add(normalize_name(t))

    if exact_matches:
        return True

    # 2) 장식기호 제거 매칭
    if comparable_name_length(decorated_normalized) >= DECORATED_NAME_MIN_CHARS:
        decorated_matches = {}
        for t in friend_texts:
            candidate = normalize_name_for_ocr_match(t)
            if candidate == decorated_normalized:
                decorated_matches[normalize_name(t)] = candidate

        if len(decorated_matches) == 1:
            raw_match = next(iter(decorated_matches.keys()))
            log(f"   -> ✅ 장식기호 제거 후 OCR 확인됨: '{raw_match}' → '{decorated_normalized}'")
            return True
        if len(decorated_matches) > 1:
            candidates = ', '.join(decorated_matches.keys())
            log(f"   -> ⚠️ 장식기호 제거 후 같은 이름 후보가 여러 개입니다: {candidates}")
            _log_ocr_failure_diagnostics(name, window_id, texts)
            return False
    else:
        log(f"   -> ⚠️ '{name}'은 이름이 짧아 장식기호 제거 OCR 비교를 건너뜁니다.")

    # 3) 부분 일치 (긴 이름에 한해 단일 후보만 허용)
    if comparable_name_length(decorated_normalized) >= SUBSTRING_MATCH_MIN_CHARS:
        substring_matches = []
        for t in friend_texts:
            candidate = normalize_name_for_ocr_match(t)
            if not candidate or candidate == decorated_normalized:
                continue
            if decorated_normalized in candidate:
                substring_matches.append(normalize_name(t))

        if len(substring_matches) == 1:
            log(
                f"   -> ✅ 부분 일치 후 OCR 확인됨: '{substring_matches[0]}' ⊇ '{decorated_normalized}'"
            )
            return True
        if len(substring_matches) > 1:
            candidates = ', '.join(substring_matches[:5])
            suffix = ' ...' if len(substring_matches) > 5 else ''
            log(f"   -> ⚠️ 부분 일치 후보가 여러 개입니다: {candidates}{suffix}")
            _log_ocr_failure_diagnostics(name, window_id, texts)
            return False

    _log_ocr_failure_diagnostics(name, window_id, texts)
    return False


def _log_ocr_failure_diagnostics(name: str, window_id: int, texts: List[str]) -> None:
    """OCR 검증 실패 시 디버깅용 후보 로그와 캡처 저장."""
    visible_texts = [normalize_name(t) for t in texts if normalize_name(t)]
    if visible_texts:
        max_preview = 30
        preview = ', '.join(visible_texts[:max_preview])
        suffix = f' ... (총 {len(visible_texts)}개)' if len(visible_texts) > max_preview else ''
        log(f"   -> ℹ️ OCR 후보: {preview}{suffix}")
    else:
        log("   -> ℹ️ OCR 후보를 읽지 못했습니다.")

    if OCR_DEBUG_SAVE:
        saved = save_window_screenshot(window_id, name)
        if saved:
            log(f"   -> 🖼 디버그 캡처 저장됨: {saved}")


def send_message_to_friend(message: str):
    """채팅방에서 메시지 전송 (메뉴 클릭 방식). 채팅방 열린 직후 창 크기 재고정."""
    pyperclip.copy(message)
    db = CHAT_DELAY_BEFORE_ENTER
    da = CHAT_DELAY_AFTER_ENTER

    # 채팅방 열기
    open_script = f'''
    tell application "KakaoTalk" to activate
    delay 0.3
    tell application "System Events"
        tell process "KakaoTalk"
            set frontmost to true
        end tell
        delay {db}
        key code 36
        delay {da}
    end tell
    '''
    run_applescript(open_script)

    # 채팅방 열린 직후 카카오톡 레이아웃 전환이 완전히 끝날 때까지 대기 후 리사이즈
    time.sleep(0.8)
    resize_kakaotalk_window(silent=True)

    # 메시지 붙여넣기 → 전송 → 채팅방 닫기
    send_script = '''
    tell application "System Events"
        tell process "KakaoTalk"
            set frontmost to true
            try
                click menu item "붙여넣기" of menu "편집" of menu bar 1
            on error
                try
                    click menu item "Paste" of menu "편집" of menu bar 1
                on error
                    try
                        click menu item "Paste" of menu "Edit" of menu bar 1
                    end try
                end try
            end try
        end tell
        delay 0.5
        key code 36
        delay 0.5
        key code 53
        delay 0.5
        key code 53
        delay 0.3
    end tell
    '''
    run_applescript(send_script)

# ============================================================
# HTML 템플릿
# ============================================================
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>카카오톡 자동 전송기</title>
    <style>
        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        .container {
            max-width: 600px;
            margin: 0 auto;
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
        }
        .header {
            background: #FEE500;
            padding: 25px;
            text-align: center;
        }
        .header h1 {
            color: #3C1E1E;
            font-size: 24px;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 10px;
        }
        .header h1::before {
            content: "💬";
            font-size: 28px;
        }
        .content {
            padding: 30px;
        }
        .section {
            background: #f8f9fa;
            border-radius: 12px;
            padding: 20px;
            margin-bottom: 20px;
        }
        .section-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
        }
        .section-header h3 {
            color: #333;
            margin: 0;
            font-size: 16px;
        }
        .filter-group {
            margin-bottom: 15px;
        }
        .filter-group:last-child {
            margin-bottom: 0;
        }
        .filter-label {
            display: inline-block;
            background: #28a745;
            color: white;
            padding: 4px 10px;
            border-radius: 15px;
            font-size: 12px;
            font-weight: bold;
            margin-bottom: 10px;
        }
        .toggle-buttons {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
        }
        .toggle-btn {
            padding: 8px 16px;
            border: 2px solid #ddd;
            border-radius: 20px;
            background: white;
            color: #555;
            font-size: 14px;
            cursor: pointer;
            transition: all 0.2s;
            user-select: none;
        }
        .toggle-btn:hover {
            border-color: #667eea;
            color: #667eea;
        }
        .toggle-btn.active {
            border-color: #667eea;
            background: #667eea;
            color: white;
        }
        .filter-actions {
            display: flex;
            gap: 8px;
        }
        .btn-filter-action {
            padding: 6px 14px;
            border: 2px solid #ddd;
            border-radius: 8px;
            background: white;
            color: #555;
            font-size: 13px;
            cursor: pointer;
            transition: all 0.2s;
        }
        .btn-filter-action:hover {
            border-color: #667eea;
            color: #667eea;
        }
        .upload-area {
            border: 2px dashed #ddd;
            border-radius: 12px;
            padding: 30px;
            text-align: center;
            margin-bottom: 20px;
            transition: all 0.3s;
            cursor: pointer;
        }
        .upload-area:hover {
            border-color: #667eea;
            background: #f8f9ff;
        }
        .upload-area.has-file {
            border-color: #28a745;
            background: #f0fff4;
        }
        .upload-area.drag-over {
            border-color: #667eea;
            background: #e8ebff;
            transform: scale(1.02);
        }
        .upload-area input[type="file"] {
            display: none;
        }
        .upload-icon {
            font-size: 48px;
            margin-bottom: 10px;
        }
        .file-name {
            color: #28a745;
            font-weight: bold;
            margin-top: 10px;
        }
        .message-section {
            background: #f8f9fa;
            border-radius: 12px;
            padding: 20px;
            margin-bottom: 20px;
        }
        .message-section h3 {
            color: #333;
            margin-bottom: 10px;
            font-size: 16px;
        }
        .message-hint {
            color: #888;
            font-size: 12px;
            margin-bottom: 10px;
        }
        .message-hint code {
            background: #e9ecef;
            padding: 2px 6px;
            border-radius: 4px;
            color: #667eea;
            font-weight: bold;
        }
        .message-textarea {
            width: 100%;
            min-height: 100px;
            padding: 12px;
            border: 2px solid #ddd;
            border-radius: 8px;
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            font-size: 14px;
            resize: vertical;
            transition: border-color 0.3s;
        }
        .message-textarea:focus {
            outline: none;
            border-color: #667eea;
        }
        .btn {
            width: 100%;
            padding: 16px;
            border: none;
            border-radius: 12px;
            font-size: 18px;
            font-weight: bold;
            cursor: pointer;
            transition: all 0.3s;
        }
        .btn-start {
            background: #FEE500;
            color: #3C1E1E;
        }
        .btn-start:hover:not(:disabled) {
            background: #E5CE00;
            transform: translateY(-2px);
        }
        .btn-start:disabled {
            background: #ccc;
            cursor: not-allowed;
        }
        .btn-stop {
            background: #dc3545;
            color: white;
        }
        .btn-stop:hover:not(:disabled) {
            background: #c82333;
            transform: translateY(-2px);
        }
        .btn-stop:disabled {
            background: #e4606d;
            cursor: not-allowed;
        }
        .log-wrapper {
            position: relative;
            margin-top: 25px;
        }
        .log-copy-btn {
            position: absolute;
            top: 10px;
            right: 14px;
            background: rgba(255,255,255,0.12);
            border: 1px solid rgba(255,255,255,0.2);
            color: #ccc;
            border-radius: 6px;
            padding: 4px 10px;
            font-size: 12px;
            cursor: pointer;
            z-index: 2;
            transition: background 0.2s;
        }
        .log-copy-btn:hover {
            background: rgba(255,255,255,0.25);
            color: #fff;
        }
        .log-area {
            background: #1e1e1e;
            border-radius: 12px;
            padding: 20px;
            padding-top: 40px;
            max-height: 900px;
            overflow-y: auto;
            font-family: 'Menlo', 'Monaco', monospace;
            font-size: 13px;
        }
        .log-area:empty::before {
            content: "로그가 여기에 표시됩니다...";
            color: #666;
        }
        .btn-pause {
            background: linear-gradient(135deg, #f5a623, #f7c948);
            color: #333;
            font-weight: 700;
        }
        .btn-pause:hover {
            background: linear-gradient(135deg, #e09500, #f5a623);
        }
        .btn-pause:disabled {
            background: #c9a84e;
            cursor: not-allowed;
            opacity: 0.6;
        }
        .log-line {
            color: #d4d4d4;
            margin: 4px 0;
            word-wrap: break-word;
        }
        .log-line.success {
            color: #4ec9b0;
        }
        .log-line.error {
            color: #f14c4c;
        }
        .log-line.warning {
            color: #cca700;
        }
        .log-line.info {
            color: #3794ff;
        }
        .status-badge {
            display: inline-block;
            padding: 6px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: bold;
            margin-bottom: 15px;
        }
        .status-idle {
            background: #e9ecef;
            color: #495057;
        }
        .status-running {
            background: #fff3cd;
            color: #856404;
            animation: pulse 1.5s infinite;
        }
        @keyframes pulse {
            0%, 100% { opacity: 1; }
            50% { opacity: 0.6; }
        }
        .footer {
            text-align: center;
            padding: 15px;
            color: #999;
            font-size: 12px;
            border-top: 1px solid #eee;
        }
        .footer .version-badge {
            color: #7c3aed;
            font-weight: 600;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>카카오톡 자동 전송기</h1>
        </div>
        <div class="content">
            <!-- 필터링 조건 선택 -->
            <div class="section">
                <div class="section-header">
                    <h3>📌 타겟 멤버 필터링 조건</h3>
                    <div class="filter-actions">
                        <button type="button" class="btn-filter-action" onclick="selectAllFilters()">전체선택</button>
                        <button type="button" class="btn-filter-action" onclick="resetFilters()">초기화</button>
                    </div>
                </div>
                <div class="filter-group">
                    <span class="filter-label">등록형태</span>
                    <div class="toggle-buttons" id="registerTypeButtons">
                        {% for rt in available_register_types %}
                        <button type="button" class="toggle-btn {% if rt in default_register_types %}active{% endif %}"
                                onclick="toggleFilter(this)" data-value="{{ rt }}">{{ rt }}</button>
                        {% endfor %}
                    </div>
                </div>
                <div class="filter-group">
                    <span class="filter-label">연령대</span>
                    <div class="toggle-buttons" id="ageGroupButtons">
                        {% for ag in available_age_groups %}
                        <button type="button" class="toggle-btn {% if ag in default_age_groups %}active{% endif %}"
                                onclick="toggleFilter(this)" data-value="{{ ag }}">{{ ag }}</button>
                        {% endfor %}
                    </div>
                </div>
            </div>

            <!-- 엑셀 파일 업로드 -->
            <div class="upload-area" id="uploadArea" 
                 onclick="document.getElementById('fileInput').click()"
                 ondragover="handleDragOver(event)"
                 ondragleave="handleDragLeave(event)"
                 ondrop="handleDrop(event)">
                <div class="upload-icon">📁</div>
                <div>엑셀 파일을 선택하거나 드래그하세요 (.xlsx)</div>
                <div class="file-name" id="fileName"></div>
                <input type="file" id="fileInput" accept=".xlsx" onchange="handleFileSelect(this)">
            </div>

            <!-- 발송 메시지 입력 -->
            <div class="message-section">
                <h3>✏️ 발송 메시지</h3>
                <div class="message-hint">
                    <code>{name}</code> 입력 시 회원 이름으로 자동 치환됩니다
                </div>
                <textarea class="message-textarea" id="messageText">{{ default_message }}</textarea>
            </div>
            
            <span class="status-badge status-idle" id="statusBadge">대기 중</span>
            
            <button class="btn btn-start" id="startBtn" disabled onclick="startSending()">
                🚀 카카오톡 전송 시작
            </button>
            <button class="btn btn-pause" id="pauseBtn" style="display:none;" onclick="togglePause()">
                ⏸ 일시정지
            </button>
            <button class="btn btn-stop" id="stopBtn" style="display:none;" onclick="stopSending()">
                ⏹ 전송 중단
            </button>
            
            <div class="log-wrapper">
                <button class="log-copy-btn" onclick="copyLog()" title="로그 복사">📋 복사</button>
                <div class="log-area" id="logArea"></div>
            </div>
        </div>
        <div class="footer">
            카카오톡 자동 전송기 <span class="version-badge">V{{ version }}</span> | 전송 중 마우스/키보드 조작 금지
        </div>
    </div>

    <script>
        let selectedFile = null;
        let eventSource = null;
        let isPaused = false;
        
        function toggleFilter(btn) {
            btn.classList.toggle('active');
        }
        
        function selectAllFilters() {
            document.querySelectorAll('#registerTypeButtons .toggle-btn, #ageGroupButtons .toggle-btn').forEach(btn => {
                btn.classList.add('active');
            });
        }
        
        function resetFilters() {
            document.querySelectorAll('#registerTypeButtons .toggle-btn, #ageGroupButtons .toggle-btn').forEach(btn => {
                btn.classList.remove('active');
            });
        }
        
        function getSelectedValues(containerId) {
            const buttons = document.querySelectorAll('#' + containerId + ' .toggle-btn.active');
            return Array.from(buttons).map(btn => btn.dataset.value);
        }
        
        function handleDragOver(event) {
            event.preventDefault();
            event.stopPropagation();
            event.currentTarget.classList.add('drag-over');
        }
        
        function handleDragLeave(event) {
            event.preventDefault();
            event.stopPropagation();
            event.currentTarget.classList.remove('drag-over');
        }
        
        function handleDrop(event) {
            event.preventDefault();
            event.stopPropagation();
            event.currentTarget.classList.remove('drag-over');
            
            const files = event.dataTransfer.files;
            if (files.length > 0) {
                const file = files[0];
                if (file.name.endsWith('.xlsx')) {
                    // FileList를 FileInput에 할당하기 위해 DataTransfer 사용
                    const dataTransfer = new DataTransfer();
                    dataTransfer.items.add(file);
                    document.getElementById('fileInput').files = dataTransfer.files;
                    handleFileSelect(document.getElementById('fileInput'));
                } else {
                    alert('엑셀 파일(.xlsx)만 업로드 가능합니다.');
                }
            }
        }
        
        function handleFileSelect(input) {
            if (input.files.length > 0) {
                selectedFile = input.files[0];
                document.getElementById('fileName').textContent = '✅ ' + selectedFile.name;
                document.getElementById('uploadArea').classList.add('has-file');
                document.getElementById('startBtn').disabled = false;
                addLog('파일 선택됨: ' + selectedFile.name, 'info');
            }
        }
        
        function addLog(message, type = '') {
            const logArea = document.getElementById('logArea');
            const line = document.createElement('div');
            line.className = 'log-line ' + type;
            line.textContent = message;
            logArea.appendChild(line);
            logArea.scrollTop = logArea.scrollHeight;
        }
        
        function startSending() {
            if (!selectedFile) return;
            
            const registerTypes = getSelectedValues('registerTypeButtons');
            const ageGroups = getSelectedValues('ageGroupButtons');
            const messageText = document.getElementById('messageText').value;
            
            if (registerTypes.length === 0) {
                alert('등록형태를 최소 1개 이상 선택해주세요.');
                return;
            }
            if (ageGroups.length === 0) {
                alert('연령대를 최소 1개 이상 선택해주세요.');
                return;
            }
            if (!messageText.trim()) {
                alert('발송 메시지를 입력해주세요.');
                return;
            }
            
            const formData = new FormData();
            formData.append('file', selectedFile);
            
            document.getElementById('startBtn').style.display = 'none';
            document.getElementById('pauseBtn').style.display = 'block';
            document.getElementById('pauseBtn').disabled = false;
            document.getElementById('pauseBtn').textContent = '⏸ 일시정지';
            isPaused = false;
            document.getElementById('stopBtn').style.display = 'block';
            document.getElementById('stopBtn').disabled = false;
            document.getElementById('stopBtn').textContent = '⏹ 전송 중단';
            document.getElementById('statusBadge').className = 'status-badge status-running';
            document.getElementById('statusBadge').textContent = '전송 중...';
            
            // 파일 업로드
            fetch('/upload', {
                method: 'POST',
                body: formData
            })
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    addLog('🚀 작업을 시작합니다...', 'info');
                    // 먼저 로그 스트림을 연 뒤 전송 시작. 초기 로그는 서버 큐에 버퍼링되어
                    // 스트림 연결 후 그대로 전달되므로(백엔드 epoch 가드로 단일 소비자 보장) 유실되지 않음.
                    startLogStream();
                    fetch('/start', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            register_types: registerTypes,
                            age_groups: ageGroups,
                            message_template: messageText
                        })
                    });
                } else {
                    addLog('❌ 오류: ' + data.error, 'error');
                    resetUI();
                }
            })
            .catch(error => {
                addLog('❌ 업로드 실패: ' + error, 'error');
                resetUI();
            });
        }
        
        function stopSending() {
            document.getElementById('stopBtn').disabled = true;
            document.getElementById('stopBtn').textContent = '⏹ 중단 중...';
            document.getElementById('pauseBtn').style.display = 'none';
            fetch('/stop', { method: 'POST' });
        }
        
        function togglePause() {
            if (!isPaused) {
                isPaused = true;
                document.getElementById('pauseBtn').textContent = '⏸ 현재 작업 완료 후 일시정지...';
                document.getElementById('pauseBtn').disabled = true;
                document.getElementById('statusBadge').textContent = '일시정지 대기...';
                fetch('/pause', { method: 'POST' });
            } else {
                isPaused = false;
                document.getElementById('pauseBtn').textContent = '⏸ 일시정지';
                document.getElementById('statusBadge').className = 'status-badge status-running';
                document.getElementById('statusBadge').textContent = '전송 중...';
                fetch('/resume', { method: 'POST' });
            }
        }
        
        function copyLog() {
            const logArea = document.getElementById('logArea');
            const lines = logArea.querySelectorAll('.log-line');
            const text = Array.from(lines).map(l => l.textContent).join('\\n');
            navigator.clipboard.writeText(text).then(() => {
                const btn = document.querySelector('.log-copy-btn');
                const orig = btn.textContent;
                btn.textContent = '✅ 복사됨';
                setTimeout(() => { btn.textContent = orig; }, 1500);
            });
        }
        
        function startLogStream() {
            if (eventSource) {
                eventSource.close();
            }
            
            eventSource = new EventSource('/logs');
            eventSource.onmessage = function(event) {
                const data = JSON.parse(event.data);
                
                if (data.type === 'log') {
                    let logType = '';
                    if (data.message.includes('✅') || data.message.includes('성공')) logType = 'success';
                    else if (data.message.includes('❌') || data.message.includes('실패')) logType = 'error';
                    else if (data.message.includes('⚠️') || data.message.includes('중단')) logType = 'warning';
                    else if (data.message.includes('🚀') || data.message.includes('📋')) logType = 'info';
                    else if (data.message.includes('⏸')) logType = 'warning';
                    else if (data.message.includes('▶️')) logType = 'info';
                    
                    addLog(data.message, logType);
                } else if (data.type === 'paused') {
                    document.getElementById('pauseBtn').textContent = '▶️ 재개';
                    document.getElementById('pauseBtn').disabled = false;
                    document.getElementById('statusBadge').className = 'status-badge status-idle';
                    document.getElementById('statusBadge').textContent = '일시정지';
                } else if (data.type === 'complete') {
                    eventSource.close();
                    resetUI();
                    if (data.stopped) {
                        alert('전송이 중단되었습니다.\\n\\n성공: ' + data.success + '/' + data.total);
                    } else if (data.failed_names && data.failed_names.length > 0) {
                        alert('완료!\\n\\n성공: ' + data.success + '/' + data.total + 
                              '\\n\\n실패한 대상자:\\n• ' + data.failed_names.join('\\n• '));
                    } else {
                        alert('완료! 모두 성공했습니다. (' + data.success + '/' + data.total + ')');
                    }
                }
            };
            
            eventSource.onerror = function() {
                eventSource.close();
                resetUI();
            };
        }
        
        function resetUI() {
            document.getElementById('startBtn').style.display = 'block';
            document.getElementById('startBtn').disabled = false;
            document.getElementById('pauseBtn').style.display = 'none';
            document.getElementById('stopBtn').style.display = 'none';
            document.getElementById('statusBadge').className = 'status-badge status-idle';
            document.getElementById('statusBadge').textContent = '대기 중';
            isPaused = false;
        }
    </script>
</body>
</html>
'''


# ============================================================
# 웹 라우트
# ============================================================
@app.route('/')
def index():
    return render_template_string(
        HTML_TEMPLATE,
        available_register_types=AVAILABLE_REGISTER_TYPES,
        available_age_groups=AVAILABLE_AGE_GROUPS,
        default_register_types=DEFAULT_REGISTER_TYPES,
        default_age_groups=DEFAULT_AGE_GROUPS,
        default_message=DEFAULT_MESSAGE_TEMPLATE,
        version=VERSION
    )


@app.route('/upload', methods=['POST'])
def upload_file():
    global current_file_path
    
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': '파일이 없습니다'})
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'error': '파일이 선택되지 않았습니다'})
        
        # 임시 파일로 저장 (항상 쓰기 가능한 OS 임시 폴더)
        temp_path = UPLOAD_TEMP_XLSX
        file.save(temp_path)
        current_file_path = temp_path
        
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/start', methods=['POST'])
def start_sending():
    global is_running, stop_requested
    global current_register_types, current_age_groups, current_message_template
    
    if is_running:
        return jsonify({'success': False, 'error': '이미 실행 중입니다'})
    
    data = request.get_json() or {}
    current_register_types = data.get('register_types', DEFAULT_REGISTER_TYPES)
    current_age_groups = data.get('age_groups', DEFAULT_AGE_GROUPS)
    current_message_template = data.get('message_template', DEFAULT_MESSAGE_TEMPLATE)
    
    is_running = True
    stop_requested = False
    pause_requested = False
    pause_event.set()
    thread = threading.Thread(target=run_sending_logic, daemon=True)
    thread.start()
    
    return jsonify({'success': True})


@app.route('/stop', methods=['POST'])
def stop_sending():
    global stop_requested
    stop_requested = True
    pause_event.set()  # 일시정지 중이면 풀어서 중단이 즉시 동작하도록
    return jsonify({'success': True})


@app.route('/pause', methods=['POST'])
def pause_sending():
    global pause_requested
    pause_requested = True
    return jsonify({'success': True})


@app.route('/resume', methods=['POST'])
def resume_sending():
    global pause_requested
    pause_requested = False
    pause_event.set()
    return jsonify({'success': True})


@app.route('/logs')
def stream_logs():
    # 연결마다 독립 큐를 만들어 구독자로 등록. 연결이 끊기면 해제한다.
    q = Queue()
    with log_subscribers_lock:
        log_subscribers.append(q)

    def generate():
        try:
            while True:
                try:
                    item = q.get(timeout=0.5)
                except Empty:
                    # 주기적 keepalive: 끊긴 연결을 쓰기 시도로 감지해 정리되게 한다.
                    yield ": keepalive\n\n"
                    continue
                yield f"data: {item}\n\n"
        finally:
            with log_subscribers_lock:
                if q in log_subscribers:
                    log_subscribers.remove(q)

    return Response(generate(), mimetype='text/event-stream')


# ============================================================
# 메시지 전송 로직
# ============================================================
def log(msg):
    """로그 메시지를 모든 스트림 구독자에게 브로드캐스트."""
    import json
    log_queue.put(json.dumps({'type': 'log', 'message': msg}))


def resize_kakaotalk_window(silent=False) -> bool:
    """면적이 가장 큰 카카오톡 창의 위치+크기를 강제 고정.
    설정 후 이중 검증(set → 대기 → 재확인 → 대기 → 최종확인)으로
    카카오톡이 뒤늦게 크기를 되돌리는 레이스 컨디션을 방지한다."""
    x, y = KAKAOTALK_WINDOW_X, KAKAOTALK_WINDOW_Y
    w, h = KAKAOTALK_WINDOW_WIDTH, KAKAOTALK_WINDOW_HEIGHT
    tolerance = 30
    max_attempts = 4

    set_script = f'''
tell application "System Events"
    tell process "KakaoTalk"
        set frontmost to true
        set wc to count of windows
        if wc is 0 then
            return "fail:no_windows"
        end if
        set maxIdx to 1
        set maxArea to 0
        repeat with i from 1 to wc
            try
                set sz to size of window i
                set ww to item 1 of sz
                set hh to item 2 of sz
                set ar to ww * hh
                if ar > maxArea then
                    set maxArea to ar
                    set maxIdx to i
                end if
            end try
        end repeat
        try
            set position of window maxIdx to {{{x}, {y}}}
            set size of window maxIdx to {{{w}, {h}}}
            return "ok:" & maxIdx
        on error errMsg
            return "fail:" & errMsg
        end try
    end tell
end tell
'''

    verify_script = f'''
tell application "System Events"
    tell process "KakaoTalk"
        set wc to count of windows
        if wc is 0 then
            return "fail:no_windows"
        end if
        set maxIdx to 1
        set maxArea to 0
        repeat with i from 1 to wc
            try
                set sz to size of window i
                set ww to item 1 of sz
                set hh to item 2 of sz
                set ar to ww * hh
                if ar > maxArea then
                    set maxArea to ar
                    set maxIdx to i
                end if
            end try
        end repeat
        set actualPos to position of window maxIdx
        set actualSize to size of window maxIdx
        set ax to item 1 of actualPos
        set ay to item 2 of actualPos
        set aw to item 1 of actualSize
        set ah to item 2 of actualSize
        return "check:" & maxIdx & ":" & ax & ":" & ay & ":" & aw & ":" & ah
    end tell
end tell
'''

    def _verify_size() -> tuple:
        """현재 창 크기를 확인하여 (성공여부, idx, aw, ah) 반환"""
        code2, out2, _ = run_applescript(verify_script)
        out2 = (out2 or "").strip()
        if out2.startswith("check:"):
            parts = out2.split(":")
            return True, parts[1], int(parts[4]), int(parts[5])
        return False, "?", 0, 0

    for attempt in range(max_attempts):
        code, out, err = run_applescript(set_script)
        out = (out or "").strip()
        err = (err or "").strip()

        if out.startswith("fail:"):
            detail = out.split(":", 1)[1] if ":" in out else (err or f"exit {code}")
            if detail == "no_windows" and attempt < max_attempts - 1:
                run_applescript('tell application "KakaoTalk" to activate')
                wait_for_kakaotalk_window(timeout=1.5)
                time.sleep(0.3)
                continue
            if not silent:
                log(f"   -> ⚠️ 카카오톡 창 크기 조정 실패: {detail}")
            return False

        if out.startswith("ok:"):
            time.sleep(0.5)
            ok, idx, aw, ah = _verify_size()
            if ok and abs(aw - w) <= tolerance and abs(ah - h) <= tolerance:
                time.sleep(0.5)
                ok2, idx2, aw2, ah2 = _verify_size()
                if ok2 and abs(aw2 - w) <= tolerance and abs(ah2 - h) <= tolerance:
                    if not silent:
                        log(f"   -> ✅ 카카오톡 창 고정 ({x},{y} → {w}×{h}), 대상 창 index {idx2}")
                    return True
                else:
                    if attempt < max_attempts - 1:
                        wait_for_kakaotalk_window(timeout=1.0)
                        time.sleep(0.3)
                        continue
            else:
                if attempt < max_attempts - 1:
                    wait_for_kakaotalk_window(timeout=1.0)
                    time.sleep(0.5)
                    continue

            if not silent:
                log(f"   -> ⚠️ 카카오톡 창 크기 불일치 (실제: {aw}×{ah}, 목표: {w}×{h}) — {max_attempts}회 재시도 후 포기")
            return False

        if not silent:
            detail = out or err or f"exit {code}"
            log(f"   -> ⚠️ 카카오톡 창 크기 조정 실패 또는 OS 제한: {detail}")
        return False
    return False


def reset_search_and_resize(silent=False) -> bool:
    """Esc·친구목록 복귀 후 창 크기 재적용 (레이아웃 전환 완료 대기 후 리사이즈)"""
    run_applescript(SCRIPT_RESET_SEARCH)
    window_id = wait_for_kakaotalk_window(timeout=3.0)
    if not window_id:
        if not silent:
            log("   -> ⚠️ 다음 검색 준비 중 카카오톡 창 복구 대기 실패")
        return False
    time.sleep(0.5)
    return resize_kakaotalk_window(silent=silent)


class StopRequestedException(Exception):
    """전송 중단 요청 예외"""
    pass


def check_stop_requested():
    """중단 요청 확인, 중단 요청 시 예외 발생"""
    if stop_requested:
        raise StopRequestedException("사용자에 의해 전송이 중단되었습니다")


def safe_sleep(duration_range, show_log=False):
    """중단 체크를 포함한 랜덤 대기
    
    Args:
        duration_range: (min, max) 튜플로 대기 시간 범위 지정
        show_log: True일 경우 대기 시간을 로그에 표시
    """
    check_stop_requested()
    delay = random.uniform(*duration_range)
    if show_log:
        log(f"   -> ⏳ {delay:.1f}초 대기...")
    elapsed = 0
    step = 0.1  # 0.1초마다 체크
    while elapsed < delay:
        time.sleep(min(step, delay - elapsed))
        elapsed += step
        check_stop_requested()


def send_message(name: str, message: str) -> bool:
    """카카오톡 메시지 전송 (OCR 검증 포함)"""
    try:
        check_stop_requested()
        
        # 1. 카카오톡 활성화 및 준비 확인
        window_id = ensure_kakaotalk_ready()
        check_stop_requested()
        if not window_id:
            log(f"   -> ⚠️ 카카오톡 창을 찾지 못해 복구를 시도합니다.")
            reset_search_and_resize(silent=True)
            window_id = ensure_kakaotalk_ready()
            check_stop_requested()
            if not window_id:
                log(f"   -> ❌ 카카오톡 창을 찾을 수 없습니다.")
                return False

        resize_kakaotalk_window()

        # 2~3. 친구 검색 + OCR 검증 (검색 화면이 안 떴거나 OCR 타이밍 문제일 수 있어 1회 재시도)
        verified = False
        for attempt in range(MAX_SEARCH_OCR_ATTEMPTS):
            check_stop_requested()
            suffix = f" (재시도 {attempt}/{MAX_SEARCH_OCR_ATTEMPTS - 1})" if attempt else ""
            log(f"   -> 📋 '{name}' 검색 중...{suffix}")
            search_ok = search_friend(name)
            if not search_ok:
                log("   -> ⚠️ 검색창 입력 검증에 실패했지만 OCR로 추가 확인을 시도합니다.")
            safe_sleep((1.0, 2.0))  # 검색 결과 로딩 대기 (랜덤, 중단 체크 포함)
            # 검색 UI 전환으로 창이 줄어든 뒤 OCR 전에 다시 고정
            resize_kakaotalk_window(silent=True)

            check_stop_requested()
            window_id = ensure_kakaotalk_ready()
            check_stop_requested()
            if not window_id:
                log(f"   -> ⚠️ OCR 전 카카오톡 창을 찾지 못해 복구를 시도합니다.")
                reset_search_and_resize(silent=True)
                window_id = ensure_kakaotalk_ready()
                check_stop_requested()
                if not window_id:
                    log(f"   -> ❌ 카카오톡 창을 찾을 수 없습니다.")
                    return False

            # 1순위: 접근성(AX) API로 검증 (정확, 부수효과 없음)
            check_stop_requested()
            if verify_friend_by_ax(name):
                verified = True
                break

            # 2순위: AX 미확인 시 기존 OCR로 폴백 (퇴보 없이 안전망 유지)
            log(f"   -> 🔍 OCR 검증 중...")
            check_stop_requested()
            if verify_friend_by_ocr(name, window_id):
                verified = True
                break

            if attempt < MAX_SEARCH_OCR_ATTEMPTS - 1:
                log("   -> ↻ 검색을 한 번 더 시도합니다.")
                reset_search_and_resize(silent=True)
                safe_sleep((0.5, 1.0))

        if not verified:
            log(f"   -> ❌ '{name}' 친구를 찾을 수 없습니다. (AX·OCR 검증 실패)")
            return False

        log(f"   -> ✅ '{name}' 친구 확인됨!")

        # 4. 메시지 전송
        check_stop_requested()
        log(f"   -> 📤 메시지 전송 중...")
        send_message_to_friend(message)
        safe_sleep((0.3, 0.8))  # 전송 후 대기 (랜덤, 중단 체크 포함)
        log(f"   -> ✅ 전송 완료!")
        return True
        
    except StopRequestedException:
        log(f"   -> ⏹ 전송 중단됨")
        raise  # 상위로 전파하여 즉시 중단
    except Exception as e:
        log(f"   -> ❌ 오류 발생: {e}")
        return False
    finally:
        # 성공/실패 관계없이 다음 검색을 위해 검색창 초기화 (중단 요청이 아닌 경우에만)
        if not stop_requested:
            try:
                reset_search_and_resize(silent=True)
                # 중단 요청이 없을 때만 대기
                if not stop_requested:
                    time.sleep(random.uniform(0.2, 0.5))
            except:
                pass


def run_sending_logic():
    """메인 전송 로직"""
    global is_running, stop_requested, pause_requested
    import json
    
    try:
        df = pd.read_excel(current_file_path)
        log(f"📊 전체 {len(df)}명 로드됨")
        
        # 선택된 필터로 타겟 멤버 필터링
        register_types = current_register_types or DEFAULT_REGISTER_TYPES
        age_groups = current_age_groups or DEFAULT_AGE_GROUPS
        message_template = current_message_template or DEFAULT_MESSAGE_TEMPLATE
        
        log(f"📌 필터 - 등록형태: {', '.join(register_types)} / 연령대: {', '.join(age_groups)}")
        
        target_df = df[
            (df['등록형태'].isin(register_types)) &
            (df['연령'].isin(age_groups))
        ]
        
        count = len(target_df)
        log(f"✅ 타겟 멤버 {count}명 필터링됨")
        if count > 0:
            filtered_names = [str(x) for x in target_df['이름'].tolist()]
            log(f"   📋 필터링된 멤버: {', '.join(filtered_names)}")
        
        if count == 0:
            log("⚠️ 타겟 멤버가 없습니다.")
            log_queue.put(json.dumps({
                'type': 'complete',
                'success': 0,
                'total': 0,
                'failed_names': [],
                'stopped': False
            }))
            return

        # 동명이인 체크 (공백 정규화 후 비교)
        normalized_names = target_df['이름'].apply(normalize_name)
        duplicated_names = normalized_names[normalized_names.duplicated(keep=False)].unique().tolist()
        if duplicated_names:
            names_str = ', '.join(duplicated_names)
            log(f"❌ 동명이인 오류: 전송 대상에 같은 이름이 존재합니다 → {names_str}")
            log("🚫 동명이인이 있을 경우 카카오톡 검색 오류가 발생할 수 있어 전송을 중단합니다.")
            log("📋 엑셀 파일에서 해당 이름의 중복 여부를 확인한 후 다시 시도해주세요.")
            log_queue.put(json.dumps({
                'type': 'complete',
                'success': 0,
                'total': count,
                'failed_names': [],
                'stopped': True
            }))
            return

        # 이모티콘/감지 기호 포함 이름 차단 (OCR·환경 이슈 방지)
        emoji_rows = target_df['이름'].apply(name_contains_emoji_or_symbol)
        if emoji_rows.any():
            bad_names = target_df.loc[emoji_rows, '이름'].unique().tolist()
            names_str = ', '.join(str(x) for x in bad_names)
            log(f"❌ 이모티콘 오류: 이름에 이모티콘 또는 지원하지 않는 기호가 포함된 행이 있습니다 → {names_str}")
            log("🚫 이모티콘이 포함된 이름은 자동 전송에서 지원하지 않아 실행을 중단합니다.")
            log("📋 엑셀에서 이모티콘을 제거하거나, 카카오톡 표시 이름과 맞게 텍스트만 남긴 후 다시 시도해주세요.")
            log_queue.put(json.dumps({
                'type': 'complete',
                'success': 0,
                'total': count,
                'failed_names': [],
                'stopped': True
            }))
            return

        # 전송 시작 전 카카오톡 사전 준비
        log("💬 카카오톡 준비 중...")
        
        window_id = None
        max_prepare_retries = 5  # 최대 5회 시도 (카카오톡이 꺼져있을 경우 시작까지 시간 필요)
        
        for attempt in range(max_prepare_retries):
            run_applescript('''
            tell application "KakaoTalk"
                activate
                delay 2.0
            end tell
            tell application "System Events"
                tell process "KakaoTalk"
                    set frontmost to true
                    if (count of windows) is 0 then
                        keystroke "n" using command down
                        delay 1.0
                    end if
                end tell
            end tell
            ''')
            time.sleep(2)
            
            window_id = get_kakaotalk_window_id()
            if window_id:
                break
            
            log(f"   -> 카카오톡 창 대기 중... ({attempt + 1}/{max_prepare_retries})")
            time.sleep(2)
        
        if not window_id:
            log("❌ 카카오톡을 찾을 수 없습니다. 카카오톡이 설치되어 있고 로그인되어 있는지 확인해주세요.")
            log_queue.put(json.dumps({
                'type': 'complete',
                'success': 0,
                'total': count,
                'failed_names': [],
                'stopped': False
            }))
            return

        # 친구 목록 복귀 + 창 크기 (면적 최대 창 기준)
        reset_search_and_resize()
        time.sleep(1)
        log("✅ 카카오톡 준비 완료!")
        
        success_count = 0
        failed_names = []
        stopped = False
        
        for i, (_, row) in enumerate(target_df.iterrows()):
            # 일시정지 체크 (현재 대상자 시작 전에 확인)
            if pause_requested:
                log(f"⏸ 일시정지됨 — 재개 버튼을 누르면 [{i + 1}/{count}]번째부터 이어서 전송합니다.")
                log_queue.put(json.dumps({'type': 'paused'}))
                pause_event.clear()
                pause_event.wait()
                if stop_requested:
                    log(f"\n⚠️ 사용자에 의해 전송이 중단되었습니다. ({i}/{count} 처리됨)")
                    stopped = True
                    break
                log(f"▶️ 전송 재개!")

            # 중단 요청 확인
            check_stop_requested()
            
            name = row['이름']
            message = message_template.format(name=name)
            
            log(f"[{i + 1}/{count}] {name}님 처리 중...")
            
            try:
                if send_message(name, message):
                    success_count += 1
                else:
                    failed_names.append(name)
            except StopRequestedException:
                log(f"\n⚠️ 사용자에 의해 전송이 중단되었습니다. ({i}/{count} 처리됨)")
                stopped = True
                break
            
            # 매크로 탐지 방지를 위한 랜덤 대기 (1~3초, 중단 체크 포함)
            check_stop_requested()
            try:
                safe_sleep((1.0, 3.0), show_log=True)
            except StopRequestedException:
                log(f"\n⚠️ 사용자에 의해 전송이 중단되었습니다. ({i + 1}/{count} 처리됨)")
                stopped = True
                break
        
        log(f"\n{'='*40}")
        if stopped:
            log(f"⚠️ 중단됨! (성공: {success_count}/{count})")
        else:
            log(f"🎉 완료! (성공: {success_count}/{count})")
        
        if failed_names:
            log(f"\n❌ 실패한 타겟 멤버 ({len(failed_names)}명):")
            for name in failed_names:
                log(f"   • {name}")
        
        log(f"{'='*40}")
        
        log_queue.put(json.dumps({
            'type': 'complete',
            'success': success_count,
            'total': count,
            'failed_names': failed_names,
            'stopped': stopped
        }))
        
    except Exception as e:
        log(f"❌ 에러 발생: {e}")
        log_queue.put(json.dumps({
            'type': 'complete',
            'success': 0,
            'total': 0,
            'failed_names': [],
            'stopped': False
        }))
    finally:
        is_running = False
        stop_requested = False
        pause_requested = False
        pause_event.set()


# ============================================================
# 메인 실행
# ============================================================
if __name__ == "__main__":
    port = 5050
    print(f"\n{'='*50}")
    print(f"  카카오톡 자동 전송기 (웹 버전)")
    print(f"{'='*50}")
    print(f"\n  브라우저에서 열림: http://localhost:{port}")
    print(f"  종료: Ctrl+C\n")
    
    # 브라우저 자동 열기
    webbrowser.open(f'http://localhost:{port}')
    
    # Flask 서버 시작
    app.run(host='127.0.0.1', port=port, debug=True, threaded=True, use_reloader=False)
