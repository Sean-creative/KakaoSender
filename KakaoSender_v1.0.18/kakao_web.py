"""
카카오톡 자동 메시지 전송기 (macOS) - 웹 버전
- 브라우저 기반 인터페이스 (tkinter 사용 안함)
- Flask 웹 서버 사용
"""

import os
import sys
import re
import csv
import unicodedata
import tempfile
import subprocess
import threading
import time
import random
import webbrowser
from contextlib import contextmanager
from queue import Queue, Empty
from datetime import datetime
from typing import Optional, List

import pandas as pd
import pyperclip
from flask import Flask, render_template_string, request, jsonify, Response

# 접근성(AX) API — 친구 이름을 정확한 문자열로 읽고/입력하기 위한 경로.
# 카카오톡 조작에 필요한 '손쉬운 사용(접근성)' 권한과 동일한 권한을 사용한다.
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

# 접근성 권한 '요청 프롬프트' API — 권한이 없을 때 macOS 표준 권한 요청 다이얼로그를
# 띄워 사용자를 '시스템 설정 > 손쉬운 사용' 패널로 바로 유도한다.
# (구버전 pyobjc 등에서 심볼이 없을 수 있어 별도 try로 감싸 기본 AX 경로를 막지 않는다.)
try:
    from ApplicationServices import (
        AXIsProcessTrustedWithOptions,
        kAXTrustedCheckOptionPrompt,
    )
    AX_PROMPT_AVAILABLE = AX_AVAILABLE
except Exception:
    AX_PROMPT_AVAILABLE = False

# 접근성(AX) '쓰기' 경로 — 검색어/메시지 입력과 전송을 키보드 없이 처리하기 위한 API.
# 사용 불가 시 자동으로 기존 AppleScript(키 입력) 방식으로 폴백한다.
try:
    from ApplicationServices import (
        AXUIElementSetAttributeValue,
        AXUIElementIsAttributeSettable,
    )
    AX_WRITE_AVAILABLE = AX_AVAILABLE
except Exception:
    AX_WRITE_AVAILABLE = False

# ============================================================
# 설정
# ============================================================
VERSION = "1.0.18"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# 카카오톡 자동화 튜닝 (다른 맥에서 검색/채팅 진입 안정화)
SEARCH_RESULT_DOWN_ARROW_COUNT = 2
MAX_SEARCH_ATTEMPTS = 2
SEARCH_INPUT_VERIFY_ATTEMPTS = 2
# 친구 검증은 접근성(AX) API로만 수행한다. (정확한 문자열 비교 → 오발송 방지)
USE_AX_VERIFICATION = True
# 검색어/메시지 입력과 전송을 접근성(AX) API로 처리(키보드 비의존). 실패 시 키 입력으로 폴백.
# AX 입력은 카카오톡이 최전면이 아니어도 동작해, 전송 중 다른 작업으로 포커스가 바뀌어도 안전하다.
USE_AX_INPUT = True
# 카카오톡 번들 식별자 (AX 앱 핸들 탐색용)
KAKAO_BUNDLE_IDS = ('com.kakao.KakaoTalkMac', 'com.kakao.KakaoTalk')
# 검색 결과에서 친구 이름이 담기는 AXStaticText 의 identifier
AX_DISPLAY_NAME_ID = 'Display Name'
CHAT_DELAY_BEFORE_ENTER = 0.55
CHAT_DELAY_AFTER_ENTER = 1.35
# 업로드 파일은 스크립트 폴더가 아닌 시스템 임시 디렉터리에 저장 (폴더명 공백·복사본 경로 등으로 인한 ENOENT 방지)
UPLOAD_TEMP_XLSX = os.path.join(tempfile.gettempdir(), f'kakao_sender_upload_{os.getpid()}.xlsx')

# ============================================================
# 단계별 소요시간 계측 (B단계) — 패스트 모드 설계를 위한 실측 도구.
# 사용자 로그(log_queue)와 분리된 채널로 기록하고, 전송 종료 시 CSV로 덤프 +
# 요약 몇 줄만 로그에 남긴다. 계측 오버헤드는 perf_counter 호출(마이크로초급)뿐.
# 비활성화하려면 TIMING_ENABLED = False.
# ============================================================
TIMING_ENABLED = True
TIMING_DIR = os.path.join(tempfile.gettempdir(), 'kakao_sender_timing')
_timing_records = []  # list[tuple(idx, name, stage, seconds)]
_timing_lock = threading.Lock()
_timing_ctx = {'idx': None, 'name': None}
# 단계 출력 순서(요약 표 정렬용)
TIMING_STAGE_ORDER = [
    'ensure_ready', 'search', 'search_result_wait', 'verify_ax',
    'open_chat', 'layout_wait', 'image_send', 'ax_input_send', 'close_chat',
    'send_total', 'post_send_wait', 'reset_search', 'person_total',
]


def set_timing_context(idx, name):
    """이번에 처리할 대상자 정보를 계측 컨텍스트에 설정 (전송 스레드 단일 → 전역 안전)."""
    _timing_ctx['idx'] = idx
    _timing_ctx['name'] = name


@contextmanager
def time_stage(stage: str):
    """with 블록의 소요시간을 단계명과 함께 기록. TIMING_ENABLED=False면 무동작."""
    if not TIMING_ENABLED:
        yield
        return
    t0 = time.perf_counter()
    try:
        yield
    finally:
        dt = time.perf_counter() - t0
        with _timing_lock:
            _timing_records.append((_timing_ctx['idx'], _timing_ctx['name'], stage, dt))


def reset_timing():
    """새 전송 시작 시 이전 계측 기록을 비운다."""
    with _timing_lock:
        _timing_records.clear()


def dump_timing_summary():
    """누적된 계측 기록을 CSV로 저장하고, 단계별 요약을 사용자 로그로 남긴다.
    반환: 저장된 CSV 경로(또는 None)."""
    if not TIMING_ENABLED:
        return None
    with _timing_lock:
        records = list(_timing_records)
    if not records:
        return None

    # 단계별 집계
    by_stage = {}
    for _idx, _name, stage, sec in records:
        by_stage.setdefault(stage, []).append(sec)

    def stat(vals):
        n = len(vals)
        avg = sum(vals) / n
        s = sorted(vals)
        mid = s[n // 2] if n % 2 else (s[n // 2 - 1] + s[n // 2]) / 2
        return n, avg, mid, min(vals), max(vals)

    person_totals = by_stage.get('person_total', [])
    n_people = len(person_totals)

    # CSV 덤프 (raw + 요약)
    csv_path = None
    try:
        os.makedirs(TIMING_DIR, exist_ok=True)
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        csv_path = os.path.join(TIMING_DIR, f'timing_{ts}.csv')
        with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
            w = csv.writer(f)
            w.writerow(['idx', 'name', 'stage', 'seconds'])
            for idx, name, stage, sec in records:
                w.writerow([idx, name, stage, f'{sec:.4f}'])
    except Exception:
        csv_path = None

    # 요약 로그 (단계 순서대로)
    ordered = [s for s in TIMING_STAGE_ORDER if s in by_stage]
    ordered += [s for s in by_stage if s not in TIMING_STAGE_ORDER]
    log("⏱ 단계별 소요시간(계측) — 평균/중앙값/최소/최대 (초)")
    for stage in ordered:
        if stage == 'person_total':
            continue
        n, avg, mid, lo, hi = stat(by_stage[stage])
        log(f"   • {stage:<18} avg {avg:5.2f} | med {mid:5.2f} | {lo:5.2f}~{hi:5.2f} (n={n})")
    if person_totals:
        n, avg, mid, lo, hi = stat(person_totals)
        log(f"   ▷ 1인당 합계        avg {avg:5.2f} | med {mid:5.2f} | {lo:5.2f}~{hi:5.2f} (n={n})")
        log(f"   ▷ {n_people}명 추정 총시간 ≈ {avg * n_people:.0f}초 (≈ {avg * n_people / 60:.1f}분)")
    if csv_path:
        log(f"   🗂 상세 CSV: {csv_path}")
    return csv_path

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
current_dry_run = False  # 모의 전송(테스트) 모드: 실제 메시지 발송을 생략
# 패스트 모드: 고정 대기를 상태 폴링으로 바꾸고 AppleScript delay를 최소화하며,
# 매크로 탐지 방지용 대기(대상 간/전송 후)를 제거한다.
# ⚠️ 빠른 연속 발송은 카카오톡 스팸/매크로 탐지로 계정이 제한될 수 있다.
current_fast_mode = False
# 이미지 첨부(선택, 1장). current_image_path가 있으면 텍스트와 함께 전송한다.
# current_image_order: 'image_first'(기본, 사진 먼저) | 'text_first'(텍스트 먼저)
current_image_path = None
current_image_order = 'image_first'


def fast_delay(normal: float, fast: float) -> float:
    """현재 모드에 따른 지연값(초). 패스트 모드면 fast, 아니면 normal을 반환.
    AppleScript의 'delay {fast_delay(...)}'와 Python time.sleep 양쪽에 공용으로 쓴다."""
    return fast if current_fast_mode else normal


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


def _reliable_copy(text: str, retries: int = 3, verify_delay: float = 0.05) -> bool:
    """클립보드에 text를 복사하고 즉시 검증한다.

    macOS 보편적 클립보드(Handoff)가 켜져 있어도 복사 직후 값을 확인해
    다른 기기가 덮어쓴 경우 재시도함으로써 오염 위험을 최소화한다.
    반환값: 최종적으로 클립보드 내용이 text와 일치하면 True.
    """
    for attempt in range(retries):
        pyperclip.copy(text)
        time.sleep(verify_delay)
        if pyperclip.paste() == text:
            return True
        if attempt < retries - 1:
            log(f"   -> ⚠️ 클립보드 검증 실패 (재시도 {attempt + 1}/{retries - 1})...")
    log("   -> ⚠️ 클립보드 내용 확인 불가 — Handoff 비활성화를 권장합니다.")
    return False


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


def normalize_name_for_match(name: str) -> str:
    """이름 매칭용 정규화: 이모티콘/장식 기호 제거 + 공백 정리.
    (카카오톡 표시 이름에 붙은 이모티콘 '홍길동🍪' 등을 떼고 비교하기 위함)"""
    text = EMOJI_PATTERN.sub('', str(name))
    return normalize_name(text)


def name_contains_emoji_or_symbol(name: str) -> bool:
    """이름에 이모티콘/감지용 특수기호가 포함되는지 (전송 차단 판별용)"""
    return bool(EMOJI_PATTERN.search(str(name)))


def canonicalize_name(name: str) -> str:
    """정확 일치 비교용 정규화: NFC 정규화 + 변형 선택자 제거 + 공백 정리.

    같은 글자의 다른 표기(예: ❤ vs ❤️, 한글 조합형/분해형)를 동일하게 보되,
    서로 다른 사람을 합치지는 않는다 → 오발송 위험을 늘리지 않으면서
    이모티콘 표현형 차이로 인한 '미발송'만 줄인다.
    """
    text = unicodedata.normalize('NFC', str(name))
    text = re.sub(r'[\uFE00-\uFE0F]', '', text)  # 변형 선택자(presentation) 제거
    return normalize_name(text)


def name_for_search(name: str) -> str:
    """카카오톡 검색창에 넣을 검색어.

    이모티콘이 포함된 이름은 '텍스트만'으로 검색해 카카오톡 필터 신뢰성을 높인다.
    (최종 친구 식별은 이모티콘까지 포함한 정확 일치로 별도 수행 → 오발송 방지)
    """
    if name_contains_emoji_or_symbol(name):
        stripped = normalize_name_for_match(name)
        if stripped:
            return stripped
    return name


# ============================================================
# AppleScript 명령어
# ============================================================
SCRIPT_ACTIVATE = '''
tell application "KakaoTalk" to activate
'''

# 검색창 초기화 (다음 검색을 위해)
def _reset_search_script() -> str:
    """다음 검색 준비 스크립트(모드별 delay): Esc 3회로 레이어 닫고 Cmd+1로 친구 목록 복귀."""
    return f'''
tell application "KakaoTalk" to activate
delay {fast_delay(0.3, 0.15)}
tell application "System Events"
    tell process "KakaoTalk"
        set frontmost to true
    end tell
    delay {fast_delay(0.2, 0.1)}

    -- 1. Esc 3회 (채팅창/검색창/알림 등 모든 레이어 닫기)
    key code 53
    delay {fast_delay(0.3, 0.1)}
    key code 53
    delay {fast_delay(0.3, 0.1)}
    key code 53
    delay {fast_delay(0.3, 0.1)}

    -- 2. 친구 목록으로 이동 (Cmd+1)
    keystroke "1" using command down
    delay {fast_delay(0.5, 0.2)}
end tell
'''


# ============================================================
# 카카오톡 활성화 헬퍼 (검증·전송은 AX 기반 — Quartz 창 탐지/캡처는 사용하지 않음)
# ============================================================
def is_kakaotalk_running() -> bool:
    """카카오톡 앱이 실행 중인지 AX(NSWorkspace)로 확인. 확인 불가 환경이면 True(낙관)."""
    if not AX_AVAILABLE:
        return True
    try:
        return _ax_get_kakao_app_element() is not None
    except Exception:
        return True


def ensure_kakaotalk_ready() -> bool:
    """카카오톡을 활성화하고 최전면으로 올린다(창이 없으면 새 창 시도). 준비되면 True.

    검증·전송은 AX(NSWorkspace 기반)라 창 ID(Quartz)가 필요 없다. 따라서 Quartz 창 탐지는
    하지 않고, AppleScript로 활성화만 한 뒤 앱 실행 여부를 AX로 확인한다."""
    script = f'''
    tell application "KakaoTalk"
        activate
        delay {fast_delay(0.5, 0.15)}
    end tell
    tell application "System Events"
        tell process "KakaoTalk"
            set frontmost to true
            -- 창이 없으면 새 창 열기 시도
            if (count of windows) is 0 then
                keystroke "n" using command down
                delay {fast_delay(0.5, 0.3)}
            end if
        end tell
    end tell
    '''
    run_applescript(script)
    time.sleep(fast_delay(0.5, 0.1))
    return is_kakaotalk_running()


def _paste_into_search(name: str, use_keystroke: bool) -> None:
    """검색창 열고 검색어를 입력.

    1순위(use_keystroke=False): 검색창을 열고 한 글자로 필드를 띄운 뒤 AX로 이름을 직접 써넣는다.
                                 (키보드 비의존, 이모티콘 이름도 정확, 포커스 영향 없음)
    2순위(use_keystroke=True 또는 AX 실패): 기존 방식(클립보드 + 메뉴 클릭/Cmd+V).
    """
    # AX 경로: 검색창 열기 → 필드 노출용 1글자 입력 → AX로 전체 값 덮어쓰기
    if AX_WRITE_AVAILABLE and USE_AX_INPUT and not use_keystroke:
        prime_script = f'''
        tell application "KakaoTalk" to activate
        delay {fast_delay(0.3, 0.15)}
        tell application "System Events"
            tell process "KakaoTalk"
                set frontmost to true
            end tell
            delay {fast_delay(0.2, 0.1)}
            key code 53
            delay {fast_delay(0.2, 0.1)}
            keystroke "1" using command down
            delay {fast_delay(0.3, 0.15)}
            key code 3 using command down
            delay {fast_delay(0.4, 0.25)}
            keystroke "."
            delay {fast_delay(0.3, 0.15)}
        end tell
        '''
        run_applescript(prime_script)
        time.sleep(fast_delay(0.2, 0.1))
        if _ax_write_search(name):
            return
        # AX 실패 → 아래 키 입력 폴백 (Cmd+A + delete로 프라임 문자 '.'까지 함께 정리됨)

    _reliable_copy(name)
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
        delay 0.2

        -- 채팅창/검색창 닫기
        key code 53
        delay 0.2

        -- 친구 목록으로 이동 (Cmd+1)
        keystroke "1" using command down
        delay 0.3

        -- 검색창 열기 (Cmd+F)
        key code 3 using command down
        delay 0.4

        -- 기존 검색어 전체 선택 + 삭제
        key code 0 using command down
        delay 0.15
        key code 51
        delay 0.2

        -- 붙여넣기 (메뉴 클릭 또는 Cmd+V)
        {paste_block}
        delay 0.5
    end tell
    '''
    run_applescript(script)


def _read_search_field_text() -> str:
    """검색창(AXSearchField)에 들어 있는 텍스트를 접근성(AX) API로 직접 읽는다.
    부수효과 없이 정확하다. 읽지 못하면 빈 문자열을 반환한다.
    (클립보드 경유 폴백은 제거 — AX 읽기가 신뢰 가능하고 클립보드 오염 위험을 없앤다.)"""
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
        # 검색 결과가 로드돼야 아래화살표가 결과를 선택하고 Enter로 채팅방이 열린다.
        # 이 대기는 채팅방 열기 신뢰성에 직결되므로 패스트 모드라도 줄이지 않는다.
        time.sleep(0.3)
        actual = normalize_name(_read_search_field_text())
        if actual == expected:
            verified = True
            break
        method = "Cmd+V" if use_keystroke else "메뉴 클릭"
        log(f"   -> ⚠️ 검색창 입력 확인 실패 ({method} 시도, 실제: '{actual[:30]}'). 재시도합니다.")

    _move_focus_to_search_results()
    return verified


# ============================================================
# 접근성(AX) 기반 친구 검증
#   - 카카오톡 검색 결과의 친구 이름을 정확한 문자열로 읽어 비교한다.
#   - 검색창(AXSearchField)과 친구행(AXStaticText id='Display Name')이
#     role/identifier로 명확히 구분되어 검색창 텍스트를 친구로 오인하지 않는다.
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


# ============================================================
# 접근성(AX) '쓰기' 헬퍼 — 키보드 없이 입력/전송
#   - 검색어/메시지 입력을 AXValue 설정으로 처리(필터링도 작동 확인됨)
#   - 전송은 입력 텍스트가 있을 때만 활성화되는 전송 버튼을 AXPress
#   - 모든 함수는 실패 시 False/None 을 반환해 호출부에서 키 입력으로 폴백한다.
# ============================================================
def _ax_set(element, attr, value) -> bool:
    """AX 속성에 값을 쓴다. 성공 시 True."""
    if not AX_WRITE_AVAILABLE or element is None:
        return False
    try:
        return AXUIElementSetAttributeValue(element, attr, value) == 0
    except Exception:
        return False


def _ax_is_settable(element, attr) -> bool:
    """AX 속성이 쓰기 가능한지 여부."""
    if not AX_WRITE_AVAILABLE or element is None:
        return False
    try:
        err, settable = AXUIElementIsAttributeSettable(element, attr, None)
        return err == 0 and bool(settable)
    except Exception:
        return False


def _ax_get_search_field_element(window, max_depth: int = 30):
    """검색창(AXSearchField) 요소 자체를 반환. 없으면 None."""
    for el in _ax_walk(window, max_depth):
        if _ax_copy(el, "AXSubrole") == "AXSearchField":
            return el
    return None


def _ax_get_message_input(window, max_depth: int = 35):
    """채팅방 메시지 입력창(AXTextArea, 설명='메시지 입력')을 반환. 없으면 None.
    채팅 말풍선도 AXTextArea라서 설명/placeholder로 입력창만 정확히 식별한다."""
    for el in _ax_walk(window, max_depth):
        if _ax_copy(el, "AXRole") != "AXTextArea":
            continue
        desc = _ax_copy(el, "AXDescription")
        placeholder = _ax_copy(el, "AXPlaceholderValue")
        if (desc and "메시지" in str(desc)) or (placeholder and "메시지" in str(placeholder)):
            return el
    return None


def _ax_write_search(name: str) -> bool:
    """AX로 검색창에 name을 직접 써넣는다(필터링도 이 값으로 작동).
    검색창은 비어 있으면 AX 트리에 안 보이므로, 호출 전 최소 1글자가 입력돼 있어야 한다.
    성공 시 True."""
    if not (AX_WRITE_AVAILABLE and USE_AX_INPUT):
        return False
    try:
        if not AXIsProcessTrusted():
            return False
        window = _ax_get_main_window(_ax_get_kakao_app_element())
        if window is None:
            return False
        field = _ax_get_search_field_element(window)
        if field is None or not _ax_is_settable(field, "AXValue"):
            return False
        if _ax_is_settable(field, "AXFocused"):
            _ax_set(field, "AXFocused", True)
            time.sleep(0.05)
        if not _ax_set(field, "AXValue", name):
            return False
        time.sleep(0.1)
        return _ax_read_search_field(window) == name
    except Exception:
        return False


def _ax_input_and_send(message: str) -> bool:
    """채팅방에서 AX로 메시지를 입력한 뒤 Enter로 전송한다. 성공 시 True.

    입력은 AX(AXValue 설정)로 처리해 클립보드 오염(Handoff)·이모티콘 문제를 피하고,
    전송만 Enter(key code 36)를 사용한다.
    (전송 버튼 AXPress는 호출이 성공해도 실제 전송을 트리거하지 못해 사용하지 않는다.)

    실패 시 입력창을 비우고 False 반환 → 호출부에서 클립보드 방식으로 폴백.
    채팅방이 이미 열려 있어야 한다."""
    if not (AX_WRITE_AVAILABLE and USE_AX_INPUT):
        return False
    msg_input = None
    try:
        if not AXIsProcessTrusted():
            return False
        window = _ax_get_main_window(_ax_get_kakao_app_element())
        if window is None:
            return False
        msg_input = _ax_get_message_input(window)
        if msg_input is None or not _ax_is_settable(msg_input, "AXValue"):
            return False

        # 메시지 입력 (AX) + 입력 검증
        if _ax_is_settable(msg_input, "AXFocused"):
            _ax_set(msg_input, "AXFocused", True)
            time.sleep(0.05)
        if not _ax_set(msg_input, "AXValue", message):
            return False
        time.sleep(0.2)
        if str(_ax_copy(msg_input, "AXValue")) != message:
            _ax_set(msg_input, "AXValue", "")
            return False

        # 전송: Enter
        run_applescript('''
        tell application "System Events"
            tell process "KakaoTalk"
                set frontmost to true
                delay 0.15
                key code 36
            end tell
        end tell
        ''')

        # 전송 확인: 입력창이 비워질 때까지 폴링한다.
        # (Enter가 실제로 전송하면 입력창이 즉시 비워짐. 끝까지 안 비워지면 미전송으로 판정 →
        #  입력창을 비우고 False를 반환해 호출부의 클립보드 폴백이 중복 없이 재시도하게 한다.)
        for _ in range(10):
            time.sleep(0.15)
            remaining = _ax_copy(msg_input, "AXValue")
            if remaining is None or str(remaining) == "":
                return True
        _ax_set(msg_input, "AXValue", "")
        return False
    except Exception as exc:
        log(f"   -> ⚠️ AX 입력/전송 예외, 클립보드 방식으로 폴백: {exc}")
        try:
            if msg_input is not None:
                _ax_set(msg_input, "AXValue", "")
        except Exception:
            pass
        return False


def _ax_wait_for_search_results(timeout: float = 1.2, interval: float = 0.05) -> bool:
    """검색 결과 행(AXStaticText)이 나타날 때까지 폴링. (패스트 모드의 고정 대기 대체)
    결과가 빨리 뜨면 즉시 True로 반환하고, 안 뜨면 timeout까지 기다린 뒤 False.
    AX를 못 쓰는 환경이면 timeout 동안 폴링하다 빠지므로 기존 고정 대기와 동급(상한)."""
    deadline = time.perf_counter() + timeout
    while time.perf_counter() < deadline:
        try:
            window = _ax_get_main_window(_ax_get_kakao_app_element())
            if window is not None and _ax_collect_result_names(window):
                return True
        except Exception:
            pass
        time.sleep(interval)
    return False


def _ax_wait_for_message_input(timeout: float = 1.5, interval: float = 0.05) -> bool:
    """채팅방 메시지 입력창(AXTextArea)이 나타날 때까지 폴링. (패스트 모드의 레이아웃 대기 대체)
    입력창이 준비된 시점에만 다음 단계로 진행하므로 미입력/오입력 실패율을 낮춘다."""
    deadline = time.perf_counter() + timeout
    while time.perf_counter() < deadline:
        try:
            window = _ax_get_main_window(_ax_get_kakao_app_element())
            if window is not None and _ax_get_message_input(window) is not None:
                return True
        except Exception:
            pass
        time.sleep(interval)
    return False


def _chat_input_ready(timeout: float) -> bool:
    """채팅방 메시지 입력창이 떴는지(=채팅방이 실제로 열렸는지) 확인.

    AX를 못 쓰는 환경(권한/라이브러리 부재)에서는 확인 수단이 없으므로 True(낙관)로 둬
    기존 동작을 유지한다. AX 사용 가능하면 입력창 등장까지 폴링한다."""
    if not (AX_AVAILABLE and USE_AX_INPUT):
        return True
    try:
        if not AXIsProcessTrusted():
            return True
    except Exception:
        return True
    return _ax_wait_for_message_input(timeout=timeout)


def verify_friend_by_ax(name: str) -> bool:
    """접근성(AX) API로 친구 검증. 확인되면 True, 아니면 False.

    검색창(AXSearchField)과 친구행(AXStaticText)이 role/identifier로 구분되므로
    검색창 텍스트를 친구로 오인할 위험이 없다.
    실패(미확인/읽기 불가) 시 친구를 찾지 못한 것으로 처리한다(오발송 방지).
    """
    if not (AX_AVAILABLE and USE_AX_VERIFICATION):
        return False

    if not AXIsProcessTrusted():
        log("   -> ⚠️ 접근성 권한이 없어 친구 검증을 할 수 없습니다. (시스템 설정 > 손쉬운 사용 권한 확인)")
        return False

    # 어떤 예외가 나더라도 크래시 없이 '미확인(False)'으로 안전하게 처리한다.
    try:
        app_element = _ax_get_kakao_app_element()
        window = _ax_get_main_window(app_element)
        if window is None:
            return False

        normalized = normalize_name(name)
        decorated_normalized = normalize_name_for_match(name)

        # 오발송 방지 가드: 검색창에 실제로 이 검색어가 들어가 있을 때만 결과를 신뢰한다.
        # (검색 필터가 안 된 채 전체 목록이 보이는 상태에서 우연히 일치해 잘못 보내는 것을 차단)
        search_value = _ax_read_search_field(window)
        if search_value is None:
            log("   -> ⚠️ 검색창을 찾지 못해 친구 검증을 보류합니다.")
            return False
        search_normalized = normalize_name(search_value)
        if search_normalized != normalized and normalize_name_for_match(search_value) != decorated_normalized:
            log(
                f"   -> ⚠️ 검색창 값('{search_normalized[:30]}')이 검색어와 달라 "
                f"친구 검증을 보류합니다."
            )
            return False

        names = _ax_collect_result_names(window)
        if not names:
            return False

        # 1) 정확 일치 (이모티콘 표현형 차이는 정규화로 흡수: ❤ == ❤️, 한글 조합/분해형)
        #    가장 흔한 경로 — 로그는 호출부의 '친구 확인됨 (AX)'로 통합.
        canon_target = canonicalize_name(name)
        if any(canonicalize_name(n) == canon_target for n in names):
            return True

        # 이모티콘이 포함된 이름은 오발송 방지를 위해 '정확 일치'만 허용한다.
        # (이모티콘을 떼면 텍스트가 같은 다른 친구에게 잘못 보내는 일을 원천 차단)
        if name_contains_emoji_or_symbol(name):
            return False

        # 2) 장식기호(이모티콘) 제거 후 일치 — 카카오톡 표시 이름에 이모티콘이 붙은 경우.
        #    (부분 일치 같은 퍼지 매칭은 AX 정확 읽기에선 불필요하고 오발송 위험이라 두지 않는다.)
        decorated_matches = {n for n in names if normalize_name_for_match(n) == decorated_normalized}
        if len(decorated_matches) == 1:
            log(f"   -> ✅ AX(장식기호 제거) 확인됨: '{next(iter(decorated_matches))}'")
            return True
        if len(decorated_matches) > 1:
            log(f"   -> ⚠️ AX 후보가 여러 개라 오발송 방지를 위해 보류: {', '.join(decorated_matches)}")
            return False

        return False
    except Exception as exc:
        log(f"   -> ⚠️ 친구 검증 중 예외 발생, 미확인으로 처리합니다: {exc}")
        return False


def get_ax_permission_state() -> dict:
    """접근성(AX) 권한 상태를 진단해 반환한다.

    state:
      - 'no_library'    : pyobjc 등 AX 라이브러리 자체가 없음(설치 문제)
      - 'no_permission' : 라이브러리는 있으나 이 프로세스에 손쉬운 사용 권한 미부여
      - 'ok'            : 친구 검증/입력 가능

    친구 검증이 AX 단독(OCR 폴백 없음)이므로, 권한이 없으면 전 대상이 실패한다.
    전송 시작 전 이 게이트로 한 번에 차단하고 명확히 안내하기 위한 함수.
    """
    if not AX_AVAILABLE:
        return {
            'state': 'no_library',
            'message': (
                '접근성 라이브러리(pyobjc)가 설치되어 있지 않습니다. '
                '실행 스크립트(카카오톡전송기.command)를 다시 실행해 패키지 설치를 완료해주세요.'
            ),
            'can_prompt': False,
        }
    if not AXIsProcessTrusted():
        return {
            'state': 'no_permission',
            'message': (
                '접근성(손쉬운 사용) 권한이 없어 카카오톡 친구를 확인할 수 없습니다.\n'
                '시스템 설정 > 개인정보 보호 및 보안 > 손쉬운 사용 에서 '
                '이 프로그램을 실행한 앱(보통 "터미널")을 켜주세요.\n'
                '권한을 켠 뒤에도 안 되면 그 앱을 완전히 종료(⌘Q)했다가 다시 실행하세요.'
            ),
            'can_prompt': AX_PROMPT_AVAILABLE,
        }
    return {'state': 'ok', 'message': '접근성 권한이 확인되었습니다.', 'can_prompt': False}


def request_ax_permission_prompt() -> dict:
    """macOS 표준 '손쉬운 사용 권한 요청' 다이얼로그를 띄운다.

    이 호출은 다이얼로그만 띄우고 '현재(아직 미부여) 상태'를 즉시 반환하므로,
    반환값으로 통과시키지 말고 호출부에서 권한 부여 후 재확인(폴링)해야 한다.
    """
    if AX_PROMPT_AVAILABLE:
        try:
            AXIsProcessTrustedWithOptions({kAXTrustedCheckOptionPrompt: True})
        except Exception:
            pass
    return get_ax_permission_state()


def _copy_image_to_clipboard(image_path: str) -> bool:
    """이미지 파일을 클립보드에 '사진'으로 복사한다. 성공 시 True.

    NSImage로 읽어 클립보드에 이미지 데이터로 올린다(파일 URL이 아니라 이미지 데이터라
    카카오톡이 '사진'으로 붙여넣는다). pyobjc/AppKit이 없으면 False.
    """
    try:
        from AppKit import NSPasteboard, NSImage
    except Exception:
        return False
    try:
        img = NSImage.alloc().initWithContentsOfFile_(image_path)
        if img is None:
            return False
        pb = NSPasteboard.generalPasteboard()
        pb.clearContents()
        return bool(pb.writeObjects_([img]))
    except Exception:
        return False


def _send_image_via_clipboard(image_path: str) -> bool:
    """열린 채팅방에 이미지 1장을 전송한다: 클립보드 복사 → '편집>붙여넣기' 메뉴 클릭 → Enter.

    Cmd+V(키 이벤트)는 채팅 연 직후 '첫 동작'일 때 입력창 포커스 레이스로 붙여넣기가 실패하는
    경우가 있었다(사진 먼저 순서에서 재현). 그래서 텍스트와 동일하게 '편집>붙여넣기' 메뉴 클릭
    방식으로 붙여넣는다 — 메뉴 클릭은 첫 동작이어도 안정적으로 입력창에 붙는다.
    붙여넣기 대기는 이미지 첨부 시간이 텍스트보다 길어 넉넉히 둔다(이미지는 1인당 1회).
    복사 실패/파일 없음 시 False(텍스트만 전송되도록 호출부에서 처리)."""
    if not image_path or not os.path.exists(image_path):
        return False
    if not _copy_image_to_clipboard(image_path):
        return False
    script = '''
    tell application "KakaoTalk" to activate
    delay 0.3
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
        delay 1.2
        key code 36
        delay 0.6
    end tell
    '''
    run_applescript(script)
    return True


def send_message_to_friend(message: str) -> bool:
    """채팅방에서 메시지(+선택적 이미지 1장) 전송. 전송을 시도했으면 True, 채팅방을
    열지 못해 전송 불가면 False.

    텍스트: AX 입력 우선(실패 시 클립보드 붙여넣기 폴백).
    이미지: 클립보드 복사 → 붙여넣기 → Enter.
    이미지가 있으면 current_image_order('image_first'=사진 먼저 / 'text_first'=텍스트 먼저)
    순서로 보낸다. (기본: 사진 먼저)
    """
    # 이미지 첨부 발송은 정상 모드에서만 안정적으로 검증됐다(패스트에선 즉시 AX 폴링·짧은
    # 대기가 렌더링 중인 카카오톡의 이미지 붙여넣기를 깨뜨림). 따라서 이미지가 있으면 이
    # '사람 처리 구간'은 패스트라도 정상 동작으로 수행한다. 패스트의 핵심 이득(대상 간/전송
    # 후 대기 제거)은 사람 사이에서 일어나므로 대량 속도는 그대로 유지된다.
    per_send_fast = current_fast_mode and not current_image_path

    # 채팅방 열기(검색결과 선택 → Enter)는 신뢰성이 최우선이라 db/activate 대기는 항상 일반값.
    db = CHAT_DELAY_BEFORE_ENTER
    da = 0.3 if per_send_fast else CHAT_DELAY_AFTER_ENTER

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

    def _open_and_check() -> bool:
        """채팅방 열기 Enter → 입력창이 뜰 때까지 대기/확인. 떴으면 True."""
        with time_stage('open_chat'):
            run_applescript(open_script)
        with time_stage('layout_wait'):
            if per_send_fast:
                return _chat_input_ready(timeout=2.0)
            time.sleep(0.8)  # 렌더링 대기
            if current_image_path:
                # 이미지 발송: 정상 모드에서 검증된 경로. 입력창 AX 폴링이 붙여넣기를
                # 방해하는 정황이 있어, 추가 AX 접근 없이 고정 대기만으로 진행한다.
                return True
            return _chat_input_ready(timeout=1.5)  # 일반 텍스트: 입력창 확인(안전망)

    # 채팅방이 안 열리는 경우(Enter 미반영 등)를 대비해 1회 재시도하고,
    # 그래도 입력창이 안 뜨면 허공 전송을 막기 위해 전송하지 않고 False를 반환한다.
    chat_ready = _open_and_check()
    if not chat_ready:
        log("   -> ↻ 채팅방이 열리지 않아 다시 시도합니다...")
        chat_ready = _open_and_check()
    if not chat_ready:
        log("   -> ⚠️ 채팅방을 열지 못해 이 대상은 전송하지 못했습니다. (입력창 미확인)")
        return False

    def _send_text(prefer_clipboard=False):
        # prefer_clipboard=True(이미지 동반)면 AX를 건너뛰고 클립보드(키보드 붙여넣기)로 보낸다.
        # AX로 텍스트를 보내면 입력창의 키보드 포커스(first responder)가 풀려, 곧이은 이미지
        # Cmd+V 붙여넣기가 입력창에 안 들어가 미리보기가 안 뜬다. 키보드 붙여넣기는 포커스를
        # 유지하므로 이미지 붙여넣기와 호환된다.
        with time_stage('ax_input_send'):
            # 1순위: AX 입력(이미지 없을 때만). 성공하면 종료.
            if not prefer_clipboard and _ax_input_and_send(message):
                return
            # 2순위(또는 이미지 동반 시 기본): 클립보드 붙여넣기 + Enter
            _reliable_copy(message)  # 붙여넣기 직전 복사 — Handoff 오염 최소화
            paste_send_script = '''
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
            end tell
            '''
            run_applescript(paste_send_script)

    def _send_image():
        # 주의: 붙여넣기→Enter는 전송 트리거이며, 텍스트처럼 '전송 완료'를 확인하지는 못한다.
        # 여기서 False는 '클립보드 준비 실패'(파일 없음/형식 불가/pyobjc 부재)를 의미한다.
        with time_stage('image_send'):
            if not _send_image_via_clipboard(current_image_path):
                log("   -> ⚠️ 이미지를 클립보드에 준비하지 못해 이미지를 건너뜁니다(텍스트만 전송). 파일/형식을 확인하세요.")

    # 텍스트(+이미지)를 순서대로 전송. 이미지가 있으면 current_image_order를 따른다.
    # 두 전송 사이엔 앞 전송이 반영될 짧은 정착 간격을 둔다(연속 전송 시 두 번째가 누락되는 것 방지).
    if current_image_path:
        # 이미지 동반 시 텍스트도 클립보드(키보드) 방식 → 입력창 포커스 유지로 이미지 붙여넣기 호환
        if current_image_order == 'text_first':
            _send_text(prefer_clipboard=True)
            time.sleep(0.5)
            _send_image()
        else:  # 기본: 사진 먼저
            _send_image()
            time.sleep(0.5)
            _send_text(prefer_clipboard=True)
    else:
        _send_text()

    # 채팅방 닫기 (Esc 2회) — AX/키 입력 경로 공통.
    # 패스트 모드는 직후 reset_search(Esc 3회 + Cmd+1)가 모든 레이어를 닫으므로 생략(중복 제거).
    if not current_fast_mode:
        close_script = '''
        tell application "System Events"
            tell process "KakaoTalk"
                set frontmost to true
                key code 53
                delay 0.4
                key code 53
                delay 0.3
            end tell
        end tell
        '''
        with time_stage('close_chat'):
            run_applescript(close_script)

    return True


def open_chat_then_close(wait_seconds: float = 1.0):
    """[모의 전송용] 선택된 검색 결과의 채팅방을 열고, 잠시 후 닫는다.

    메시지 붙여넣기/전송은 하지 않는다. 채팅방 진입·복귀 흐름만 실전과 동일하게 검증한다.
    """
    # 채팅방 열기 신뢰성을 위해 db/activate 대기는 일반값 유지(send_message_to_friend와 동일).
    db = CHAT_DELAY_BEFORE_ENTER
    da = fast_delay(CHAT_DELAY_AFTER_ENTER, 0.3)

    # 채팅방 열기 (Enter)
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

    # 채팅방 레이아웃 전환 대기 (실전과 동일). 패스트는 입력창 등장 폴링으로 대체.
    if current_fast_mode:
        _ax_wait_for_message_input(timeout=1.5)
    else:
        time.sleep(0.8)

    # 채팅방을 연 상태로 잠시 유지 (중단 요청은 즉시 반영). 패스트 모드는 짧게.
    hold = fast_delay(wait_seconds, 0.2)
    safe_sleep((hold, hold))

    # 채팅방 닫기 (Esc 2회) — 모드별 delay
    close_script = f'''
    tell application "System Events"
        tell process "KakaoTalk"
            set frontmost to true
        end tell
        delay {fast_delay(0.2, 0.1)}
        key code 53
        delay {fast_delay(0.3, 0.1)}
        key code 53
        delay {fast_delay(0.3, 0.1)}
    end tell
    '''
    run_applescript(close_script)

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
        .emoji-note {
            background: #fff8e1;
            border: 1px solid #ffe082;
            border-radius: 8px;
            padding: 10px 12px;
            margin: 12px 0;
            color: #7a5c00;
            font-size: 12px;
            line-height: 1.6;
        }
        .emoji-note b { color: #5c4400; }
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
        .dryrun-section {
            background: #fff8e1;
            border: 1px solid #ffe082;
            border-radius: 12px;
            padding: 14px 16px;
            margin-bottom: 20px;
        }
        .dryrun-label {
            display: flex;
            align-items: flex-start;
            gap: 10px;
            cursor: pointer;
            color: #5d4037;
            font-size: 14px;
            line-height: 1.5;
        }
        .dryrun-label input[type="checkbox"] {
            width: 18px;
            height: 18px;
            margin-top: 2px;
            flex-shrink: 0;
            cursor: pointer;
            accent-color: #f0a500;
        }
        .dryrun-label .dryrun-desc {
            color: #8d6e63;
            font-size: 12px;
        }
        .image-section {
            background: #eef6ff;
            border: 1px solid #bcdcff;
            border-radius: 12px;
            padding: 14px 16px;
            margin-bottom: 15px;
            text-align: left;
        }
        .image-pick {
            display: block;
            font-size: 14px;
            font-weight: bold;
            color: #1c4e80;
            margin-bottom: 8px;
        }
        .image-pick .image-opt {
            font-weight: normal;
            font-size: 12px;
            color: #5a82a8;
        }
        .image-section input[type="file"] {
            font-size: 13px;
            color: #1c4e80;
            max-width: 100%;
        }
        .image-info {
            display: flex;
            align-items: center;
            gap: 10px;
            margin-top: 10px;
        }
        .image-info .image-name {
            font-size: 13px;
            color: #1c4e80;
            word-break: break-all;
        }
        .image-info .image-remove {
            border: none;
            background: #e7f0fb;
            color: #1c4e80;
            border-radius: 8px;
            padding: 4px 10px;
            font-size: 12px;
            cursor: pointer;
            flex-shrink: 0;
        }
        .image-order {
            display: flex;
            align-items: center;
            gap: 14px;
            margin-top: 10px;
            font-size: 13px;
            color: #1c4e80;
        }
        .image-order .image-order-title {
            font-weight: bold;
        }
        .image-order .image-order-opt {
            display: inline-flex;
            align-items: center;
            gap: 5px;
            cursor: pointer;
        }
        .fastmode-section {
            background: #fdecec;
            border: 1px solid #f5b7b7;
            border-radius: 12px;
            padding: 14px 16px;
            margin-bottom: 20px;
        }
        .fastmode-label {
            display: flex;
            align-items: flex-start;
            gap: 10px;
            cursor: pointer;
            color: #842029;
            font-size: 14px;
            line-height: 1.5;
        }
        .fastmode-label input[type="checkbox"] {
            width: 18px;
            height: 18px;
            margin-top: 2px;
            flex-shrink: 0;
            cursor: pointer;
            accent-color: #d6336c;
        }
        .fastmode-label .fastmode-desc {
            color: #9c5560;
            font-size: 12px;
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
        .status-warn {
            background: #f8d7da;
            color: #842029;
        }
        @keyframes pulse {
            0%, 100% { opacity: 1; }
            50% { opacity: 0.6; }
        }
        .perm-panel {
            border: 1px solid #f5c2c7;
            background: #fff5f5;
            border-radius: 12px;
            padding: 16px 18px;
            margin-bottom: 15px;
            text-align: left;
        }
        .perm-panel .perm-title {
            font-weight: bold;
            color: #842029;
            margin-bottom: 8px;
        }
        .perm-panel .perm-msg {
            font-size: 13px;
            color: #495057;
            white-space: pre-line;
            line-height: 1.55;
            margin-bottom: 12px;
        }
        .perm-panel .perm-actions {
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
            margin-bottom: 10px;
        }
        .perm-panel .btn-perm, .perm-panel .btn-perm-sec {
            border: none;
            border-radius: 8px;
            padding: 8px 14px;
            font-size: 13px;
            font-weight: bold;
            cursor: pointer;
        }
        .perm-panel .btn-perm {
            background: #7c3aed;
            color: #fff;
        }
        .perm-panel .btn-perm-sec {
            background: #e9ecef;
            color: #495057;
        }
        .perm-panel .btn-perm-close {
            background: transparent;
            border: none;
            color: #adb5bd;
            cursor: pointer;
            font-size: 13px;
        }
        .perm-panel .perm-status {
            font-size: 12px;
            color: #856404;
            margin-bottom: 4px;
        }
        .perm-panel .perm-hint {
            font-size: 11px;
            color: #adb5bd;
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

            <!-- 이모티콘 이름 주의 안내 -->
            <div class="emoji-note">
                💡 이름에 <b>이모티콘</b>이 있어도 전송할 수 있습니다. 단, 안전을 위해 이모티콘 이름은
                <b>카카오톡 표시 이름과 글자·이모티콘까지 완전히 동일</b>할 때만 전송됩니다.<br>
                ⚠️ 이모티콘을 <b>직접 손으로 입력</b>(다른 키보드·이모지 피커)하면 같은 모양이라도 내부 코드가 달라
                전송이 안 될 수 있어요. 가급적 <b>카카오톡 이름을 복사해 붙여넣기</b> 하세요.
            </div>

            <!-- 발송 메시지 입력 -->
            <div class="message-section">
                <h3>✏️ 발송 메시지</h3>
                <div class="message-hint">
                    <code>{name}</code> 입력 시 회원 이름으로 자동 치환됩니다
                </div>
                <textarea class="message-textarea" id="messageText">{{ default_message }}</textarea>
            </div>

            <!-- 이미지 첨부 (선택, 1장) -->
            <div class="image-section">
                <label class="image-pick" for="imageInput">📷 이미지 첨부 <span class="image-opt">(선택 · 1장)</span></label>
                <input type="file" id="imageInput" accept="image/*" onchange="handleImageSelect(this)">
                <div class="image-info" id="imageInfo" style="display:none;">
                    <span class="image-name" id="imageName"></span>
                    <button type="button" class="image-remove" onclick="removeImage()">✕ 제거</button>
                </div>
                <div class="image-order" id="imageOrderRow" style="display:none;">
                    <span class="image-order-title">전송 순서</span>
                    <label class="image-order-opt"><input type="radio" name="imageOrder" value="image_first" checked> 사진 먼저</label>
                    <label class="image-order-opt"><input type="radio" name="imageOrder" value="text_first"> 텍스트 먼저</label>
                </div>
            </div>

            <!-- 모의 전송(dry-run) 옵션 -->
            <div class="dryrun-section">
                <label class="dryrun-label">
                    <input type="checkbox" id="dryRunCheck">
                    <span>🧪 모의 전송(테스트) 모드<br>
                        <span class="dryrun-desc">친구 검색·검증까지만 수행하고 <b>실제 메시지는 보내지 않습니다.</b> 대량 테스트용으로 안전합니다.</span>
                    </span>
                </label>
            </div>

            <!-- 패스트 모드 옵션 -->
            <div class="fastmode-section">
                <label class="fastmode-label">
                    <input type="checkbox" id="fastModeCheck">
                    <span>⚡ 패스트 모드<br>
                        <span class="fastmode-desc">대기를 최소화하고 단계 대기를 폴링으로 바꿔 <b>전송 시간을 크게 줄입니다.</b><br>
                        <b style="color:#b02a37;">⚠️ 대상 간 대기까지 제거하므로, 빠른 연속 발송이 카카오톡 스팸/매크로 탐지에 걸려 계정이 제한될 수 있습니다.</b> 위험을 감수할 때만 사용하세요.</span>
                    </span>
                </label>
            </div>

            <span class="status-badge status-idle" id="statusBadge">대기 중</span>

            <div class="perm-panel" id="permPanel" style="display:none;">
                <div class="perm-title">⚠️ 접근성(손쉬운 사용) 권한이 필요합니다</div>
                <div class="perm-msg" id="permMsg"></div>
                <div class="perm-actions">
                    <button type="button" class="btn-perm" id="permRequestBtn" onclick="requestAxPermission()">🔓 권한 요청 팝업 열기</button>
                    <button type="button" class="btn-perm-sec" onclick="recheckAxPermission()">🔄 다시 확인</button>
                    <button type="button" class="btn-perm-close" onclick="closePermPanel()">✕ 닫기</button>
                </div>
                <div class="perm-status" id="permStatus">권한 상태를 확인하는 중...</div>
                <div class="perm-hint">권한을 켠 뒤에도 안 되면 실행한 앱(터미널)을 완전히 종료(⌘Q)했다가 다시 실행하세요.</div>
            </div>

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
        let axPollTimer = null;
        let hasImage = false;
        
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

            // 사전 점검: 접근성 권한이 없으면 시작 자체를 막고 안내 패널을 띄운다.
            document.getElementById('startBtn').disabled = true;
            fetch('/ax_status')
                .then(r => r.json())
                .then(s => {
                    if (s.state === 'ok') {
                        closePermPanel();
                        beginSendFlow();
                    } else {
                        showPermPanel(s);
                    }
                })
                .catch(() => {
                    // 상태 확인 자체가 실패하면 일단 진행 (서버측 /start 게이트가 최종 방어선)
                    beginSendFlow();
                });
        }

        function beginSendFlow() {
            const registerTypes = getSelectedValues('registerTypeButtons');
            const ageGroups = getSelectedValues('ageGroupButtons');
            const messageText = document.getElementById('messageText').value;
            const dryRun = document.getElementById('dryRunCheck').checked;
            const fastMode = document.getElementById('fastModeCheck').checked;

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
            setImageControlsEnabled(false);  // 전송 중 이미지 교체 방지(서버도 차단)

            // 파일 업로드
            fetch('/upload', {
                method: 'POST',
                body: formData
            })
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    addLog('🚀 작업을 시작합니다...', 'info');
                    // 로그 스트림을 연결하고, 연결이 '열린 뒤'(onopen)에 전송을 시작한다.
                    // 서버는 브로드캐스트만 하고 버퍼링하지 않으므로, 구독자 등록 전에 전송을
                    // 시작하면 백엔드의 초기 로그(모의전송·패스트 모드 안내 등)가 유실된다.
                    startLogStream(function() {
                    fetch('/start', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            register_types: registerTypes,
                            age_groups: ageGroups,
                            message_template: messageText,
                            dry_run: dryRun,
                            fast_mode: fastMode,
                            attach_image: hasImage,
                            image_order: getImageOrder()
                        })
                    })
                    .then(r => r.json())
                    .then(d => {
                        // /start가 거부되면(권한 변동 등) 스트림을 닫고 안내 패널을 띄운다.
                        if (!d.success) {
                            if (eventSource) { eventSource.close(); }
                            addLog('❌ ' + (d.error || '전송을 시작할 수 없습니다.'), 'error');
                            if (d.error_code === 'no_permission' || d.error_code === 'no_library') {
                                showPermPanel({ state: d.error_code, message: d.error, can_prompt: d.can_prompt });
                            } else {
                                resetUI();
                            }
                        }
                    })
                    .catch(() => { /* 정상 진행 시 결과는 로그 스트림으로 전달됨 */ });
                    });  // ← startLogStream onopen 콜백 끝
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

        // ===== 접근성(AX) 권한 안내/요청 =====
        function showPermPanel(s) {
            resetUI();
            document.getElementById('statusBadge').className = 'status-badge status-warn';
            document.getElementById('statusBadge').textContent = '권한 필요';
            document.getElementById('permMsg').textContent = s.message || '';
            const reqBtn = document.getElementById('permRequestBtn');
            reqBtn.style.display = (s.state === 'no_permission' && s.can_prompt !== false) ? 'inline-block' : 'none';
            document.getElementById('permStatus').textContent = '⏳ 권한을 켜면 자동으로 전송이 시작됩니다.';
            document.getElementById('permPanel').style.display = 'block';
            startAxPolling();
        }

        function closePermPanel() {
            stopAxPolling();
            document.getElementById('permPanel').style.display = 'none';
            document.getElementById('startBtn').disabled = false;
            document.getElementById('statusBadge').className = 'status-badge status-idle';
            document.getElementById('statusBadge').textContent = '대기 중';
        }

        function startAxPolling() {
            stopAxPolling();
            axPollTimer = setInterval(() => {
                fetch('/ax_status')
                    .then(r => r.json())
                    .then(s => {
                        if (s.state === 'ok') {
                            stopAxPolling();
                            document.getElementById('permStatus').textContent = '✅ 권한이 확인되었습니다. 전송을 시작합니다...';
                            addLog('✅ 접근성 권한이 확인되었습니다.', 'success');
                            setTimeout(() => {
                                document.getElementById('permPanel').style.display = 'none';
                                beginSendFlow();
                            }, 800);
                        } else {
                            document.getElementById('permMsg').textContent = s.message || '';
                        }
                    })
                    .catch(() => {});
            }, 1500);
        }

        function stopAxPolling() {
            if (axPollTimer) { clearInterval(axPollTimer); axPollTimer = null; }
        }

        function requestAxPermission() {
            document.getElementById('permStatus').textContent = '시스템 권한 요청 팝업을 띄웠습니다. "시스템 설정 열기"를 눌러 권한을 켜주세요.';
            fetch('/ax_request_permission', { method: 'POST' })
                .then(r => r.json())
                .then(s => { document.getElementById('permMsg').textContent = s.message || ''; })
                .catch(() => {});
        }

        function recheckAxPermission() {
            fetch('/ax_status')
                .then(r => r.json())
                .then(s => {
                    if (s.state === 'ok') {
                        document.getElementById('permPanel').style.display = 'none';
                        stopAxPolling();
                        beginSendFlow();
                    } else {
                        document.getElementById('permMsg').textContent = s.message || '';
                        document.getElementById('permStatus').textContent = '⏳ 아직 권한이 없습니다. 설정에서 권한을 켜주세요.';
                    }
                })
                .catch(() => {});
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
        
        function handleImageSelect(input) {
            if (!input.files || !input.files[0]) return;
            const file = input.files[0];
            const fd = new FormData();
            fd.append('image', file);
            fetch('/upload_image', { method: 'POST', body: fd })
                .then(r => r.json())
                .then(d => {
                    if (d.success) {
                        hasImage = true;
                        document.getElementById('imageName').textContent = '📷 ' + (d.filename || file.name);
                        document.getElementById('imageInfo').style.display = 'flex';
                        document.getElementById('imageOrderRow').style.display = 'flex';
                    } else {
                        alert('이미지 업로드 실패: ' + (d.error || ''));
                        removeImage();
                    }
                })
                .catch(e => { alert('이미지 업로드 실패: ' + e); removeImage(); });
        }

        function removeImage() {
            hasImage = false;
            document.getElementById('imageInput').value = '';
            document.getElementById('imageInfo').style.display = 'none';
            document.getElementById('imageOrderRow').style.display = 'none';
            fetch('/clear_image', { method: 'POST' }).catch(() => {});
        }

        function getImageOrder() {
            const r = document.querySelector('input[name="imageOrder"]:checked');
            return r ? r.value : 'image_first';
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
        
        function startLogStream(onReady) {
            if (eventSource) {
                eventSource.close();
            }

            eventSource = new EventSource('/logs');
            let readyFired = false;
            eventSource.onopen = function() {
                // 구독자가 서버에 등록된 뒤 호출되므로, 여기서 전송을 시작해야 초기 로그가 유실되지 않는다.
                if (onReady && !readyFired) {
                    readyFired = true;
                    onReady();
                }
            };
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
            setImageControlsEnabled(true);  // 전송 종료 후 이미지 컨트롤 복구
        }

        function setImageControlsEnabled(enabled) {
            const inp = document.getElementById('imageInput');
            if (inp) inp.disabled = !enabled;
            const rm = document.querySelector('.image-remove');
            if (rm) rm.disabled = !enabled;
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


ALLOWED_IMAGE_EXTS = ('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.heic', '.tiff')
MAX_IMAGE_BYTES = 20 * 1024 * 1024  # 업로드 이미지 상한 (20MB)


def _looks_like_image_file(path: str) -> bool:
    """파일 헤더(매직 바이트)로 실제 이미지인지 확인. 확장자만 신뢰하지 않기 위함."""
    try:
        with open(path, 'rb') as fp:
            head = fp.read(16)
    except Exception:
        return False
    if len(head) < 12:
        return False
    return any((
        head.startswith(b'\x89PNG\r\n\x1a\n'),           # PNG
        head.startswith(b'\xff\xd8\xff'),                # JPEG
        head[:4] == b'GIF8',                             # GIF
        head.startswith(b'BM'),                          # BMP
        head[:4] == b'RIFF' and head[8:12] == b'WEBP',   # WEBP
        head[:4] in (b'II*\x00', b'MM\x00*'),            # TIFF
        head[4:8] == b'ftyp',                            # HEIC/HEIF (ftyp 박스)
    ))


def _remove_temp_image(path) -> None:
    """임시 이미지 파일을 best-effort로 삭제한다."""
    try:
        if path and os.path.exists(path):
            os.remove(path)
    except Exception:
        pass


@app.route('/upload_image', methods=['POST'])
def upload_image():
    """첨부 이미지 1장 업로드. OS 임시 폴더에 저장하고 경로를 보관한다."""
    global current_image_path
    # 전송 중에는 이미지 교체를 막는다(사람마다 다른 사진이 나가는 사고 방지).
    if is_running:
        return jsonify({'success': False, 'error': '전송 중에는 이미지를 변경할 수 없습니다. 전송을 멈춘 뒤 다시 시도하세요.'})
    try:
        if 'image' not in request.files:
            return jsonify({'success': False, 'error': '이미지가 없습니다'})
        f = request.files['image']
        if not f.filename:
            return jsonify({'success': False, 'error': '이미지가 선택되지 않았습니다'})
        ext = os.path.splitext(f.filename)[1].lower()
        if ext not in ALLOWED_IMAGE_EXTS:
            return jsonify({'success': False, 'error': f'지원하지 않는 이미지 형식입니다 ({ext or "확장자 없음"})'})
        # 크기 제한: 본문 길이로 사전 차단
        if request.content_length and request.content_length > MAX_IMAGE_BYTES:
            return jsonify({'success': False, 'error': f'이미지가 너무 큽니다 (최대 {MAX_IMAGE_BYTES // (1024 * 1024)}MB)'})
        path = os.path.join(tempfile.gettempdir(), f'kakao_sender_image_{os.getpid()}{ext}')
        f.save(path)
        # 저장 후 실제 크기/내용 재검증
        if os.path.getsize(path) > MAX_IMAGE_BYTES:
            _remove_temp_image(path)
            return jsonify({'success': False, 'error': f'이미지가 너무 큽니다 (최대 {MAX_IMAGE_BYTES // (1024 * 1024)}MB)'})
        if not _looks_like_image_file(path):
            _remove_temp_image(path)
            return jsonify({'success': False, 'error': '이미지 파일이 아니거나 손상되었습니다'})
        # 직전 이미지가 다른 경로(다른 확장자)면 정리 후 교체 — 임시파일 누적 방지
        if current_image_path and current_image_path != path:
            _remove_temp_image(current_image_path)
        current_image_path = path
        return jsonify({'success': True, 'filename': f.filename})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/clear_image', methods=['POST'])
def clear_image():
    """첨부 이미지 제거 (임시 파일도 함께 정리)."""
    global current_image_path
    if is_running:
        return jsonify({'success': False, 'error': '전송 중에는 이미지를 변경할 수 없습니다.'})
    _remove_temp_image(current_image_path)
    current_image_path = None
    return jsonify({'success': True})


@app.route('/ax_status')
def ax_status():
    """접근성(AX) 권한 상태 조회 — 프론트의 사전 점검·폴링용."""
    return jsonify(get_ax_permission_state())


@app.route('/ax_request_permission', methods=['POST'])
def ax_request_permission():
    """macOS 표준 손쉬운 사용 권한 요청 다이얼로그를 띄우고 현재 상태를 반환."""
    return jsonify(request_ax_permission_prompt())


@app.route('/start', methods=['POST'])
def start_sending():
    global is_running, stop_requested
    global current_register_types, current_age_groups, current_message_template, current_dry_run
    global current_fast_mode, current_image_path, current_image_order

    if is_running:
        return jsonify({'success': False, 'error': '이미 실행 중입니다'})

    # 사전 점검(preflight): 접근성 권한이 없으면 한 명도 보낼 수 없으므로 시작 자체를 차단한다.
    perm = get_ax_permission_state()
    if perm['state'] != 'ok':
        return jsonify({
            'success': False,
            'error_code': perm['state'],
            'error': perm['message'],
            'can_prompt': perm.get('can_prompt', False),
        })

    data = request.get_json() or {}
    current_register_types = data.get('register_types', DEFAULT_REGISTER_TYPES)
    current_age_groups = data.get('age_groups', DEFAULT_AGE_GROUPS)
    current_message_template = data.get('message_template', DEFAULT_MESSAGE_TEMPLATE)
    current_dry_run = bool(data.get('dry_run', False))
    current_fast_mode = bool(data.get('fast_mode', False))
    # 이미지 첨부: 순서 옵션 반영. attach_image=False면 이전 실행의 stale 이미지를 해제한다.
    current_image_order = 'text_first' if data.get('image_order') == 'text_first' else 'image_first'
    if not data.get('attach_image'):
        current_image_path = None

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


def reset_search(silent=False) -> bool:
    """Esc·친구목록 복귀로 다음 검색을 준비. (AppleScript Esc+Cmd+1이 복귀를 수행하며,
    AX 전송엔 창 ID가 불필요하므로 Quartz 창 확인은 하지 않는다.)"""
    run_applescript(_reset_search_script())
    time.sleep(fast_delay(0.3, 0.1))
    return True


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


def send_message(name: str, message: str, dry_run: bool = False) -> bool:
    """카카오톡 메시지 전송 (접근성(AX) 검증).

    dry_run=True이면 친구 검색·검증까지만 수행하고 실제 메시지는 보내지 않는다.
    """
    try:
        check_stop_requested()
        
        # 1. 카카오톡 활성화 및 준비 확인 (검증·전송은 AX 기반)
        with time_stage('ensure_ready'):
            ready = ensure_kakaotalk_ready()
        check_stop_requested()
        if not ready:
            log(f"   -> ⚠️ 카카오톡이 준비되지 않아 복구를 시도합니다.")
            reset_search(silent=True)
            ready = ensure_kakaotalk_ready()
            check_stop_requested()
            if not ready:
                log(f"   -> ❌ 카카오톡을 찾을 수 없습니다. (실행/로그인 확인)")
                return False

        # 2~3. 친구 검색 + AX 검증 (검색 화면이 안 떴거나 타이밍 문제일 수 있어 1회 재시도)
        #  - 이모티콘 포함 이름: 검색은 '텍스트만'으로(필터 신뢰성↑), 검증은 이모티콘까지
        #    포함한 정규화 정확 일치(AX). 검증은 접근성(AX) API로만 수행한다.
        is_emoji_name = name_contains_emoji_or_symbol(name)
        search_term = name_for_search(name)
        verified = False
        for attempt in range(MAX_SEARCH_ATTEMPTS):
            check_stop_requested()
            suffix = f" (재시도 {attempt}/{MAX_SEARCH_ATTEMPTS - 1})" if attempt else ""
            log(f"   -> 📋 검색 중...{suffix}")
            with time_stage('search'):
                search_ok = search_friend(search_term)
            if not search_ok:
                log("   -> ⚠️ 검색창 입력 검증에 실패했지만 AX로 추가 확인을 시도합니다.")
            with time_stage('search_result_wait'):
                if current_fast_mode:
                    _ax_wait_for_search_results(timeout=1.2)  # 결과 행 등장까지 폴링
                else:
                    safe_sleep((0.8, 1.2))  # 검색 결과 로딩 대기 (랜덤, 중단 체크 포함)

            # 친구 검증: 접근성(AX) API (창 크기와 무관하게 정확한 문자열 비교)
            check_stop_requested()
            with time_stage('verify_ax'):
                _ax_verified = verify_friend_by_ax(name)
            if _ax_verified:
                verified = True
                break

            if attempt < MAX_SEARCH_ATTEMPTS - 1:
                log("   -> ↻ 검색을 한 번 더 시도합니다.")
                reset_search(silent=True)
                if not current_fast_mode:
                    safe_sleep((0.5, 1.0))

        if not verified:
            if is_emoji_name:
                log(f"   -> ❌ '{name}' 친구를 찾을 수 없습니다. (이모티콘 정확 일치 실패 — 카카오톡 표시 이름과 이모티콘까지 동일해야 합니다)")
            else:
                log(f"   -> ❌ '{name}' 친구를 찾을 수 없습니다. (AX 검증 실패)")
            return False

        log(f"   -> ✅ 친구 확인됨 (AX)")

        # 4. 메시지 전송 (모의 전송 모드면 채팅방만 열었다 닫고 발송은 생략)
        check_stop_requested()
        if dry_run:
            if current_image_path:
                order_label = '텍스트 → 사진' if current_image_order == 'text_first' else '사진 → 텍스트'
                log(f"   -> 🧪 (모의 전송) 실제로는 [{order_label}] 순서로 전송됩니다 — 지금은 채팅방만 열고 닫음")
            else:
                log(f"   -> 🧪 (모의 전송) 채팅방 열고 1초 후 닫음 — 실제 메시지는 보내지 않음")
            open_chat_then_close(wait_seconds=1.0)
            return True
        with time_stage('send_total'):
            sent_ok = send_message_to_friend(message)
        if not sent_ok:
            # 채팅방을 열지 못해 전송하지 못함 → 실패로 처리(허공 전송 방지)
            return False
        with time_stage('post_send_wait'):
            if current_fast_mode:
                check_stop_requested()  # 패스트: 탐지방지 대기 제거(⚠️스팸 위험), 중단만 확인
            else:
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
                with time_stage('reset_search'):
                    reset_search(silent=True)
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
        reset_timing()  # B(계측): 이전 실행 기록 초기화
        if current_dry_run:
            log("🧪 모의 전송(테스트) 모드 — 친구 검증까지만 수행하며 실제 메시지는 전송되지 않습니다.")
        if current_fast_mode:
            log("⚡ 패스트 모드 ON — 대기 최소화로 빠르게 보냅니다. ⚠️ 빠른 연속 발송은 카카오톡 스팸/매크로 탐지로 계정이 제한될 수 있습니다.")
        if current_image_path:
            order_label = '텍스트 → 사진' if current_image_order == 'text_first' else '사진 → 텍스트'
            log(f"📷 이미지 첨부 ON — 전송 순서: {order_label} (사진 1장)")
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

        # 텍스트 없이 이모티콘/기호만으로 된 이름만 차단 (AX로도 안전한 정확 검증이 불가).
        # 텍스트가 있는 이모티콘 이름(예: '홍길동🍪')은 AX 정확 일치로 안전하게 지원한다.
        emoji_only_rows = target_df['이름'].apply(lambda x: not normalize_name_for_match(x))
        if emoji_only_rows.any():
            bad_names = target_df.loc[emoji_only_rows, '이름'].unique().tolist()
            names_str = ', '.join(str(x) for x in bad_names)
            log(f"❌ 이름 오류: 텍스트 없이 이모티콘/기호만으로 된 이름이 있습니다 → {names_str}")
            log("🚫 이런 이름은 안전한 검증이 어려워 실행을 중단합니다.")
            log("📋 엑셀에서 해당 이름에 카카오톡 표시 이름의 텍스트를 추가한 뒤 다시 시도해주세요.")
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
        
        kakao_ready = False
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

            # 창 ID(Quartz) 대신 AX로 앱 실행 여부 확인 (검증·전송은 AX 기반)
            if is_kakaotalk_running():
                kakao_ready = True
                break

            log(f"   -> 카카오톡 대기 중... ({attempt + 1}/{max_prepare_retries})")
            time.sleep(2)

        if not kakao_ready:
            log("❌ 카카오톡을 찾을 수 없습니다. 카카오톡이 설치되어 있고 로그인되어 있는지 확인해주세요.")
            log_queue.put(json.dumps({
                'type': 'complete',
                'success': 0,
                'total': count,
                'failed_names': [],
                'stopped': False
            }))
            return

        # 친구 목록으로 복귀 (다음 검색 준비)
        reset_search()
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
            set_timing_context(i + 1, name)  # B(계측): 이번 대상자 컨텍스트

            try:
                with time_stage('person_total'):
                    sent_ok = send_message(name, message, dry_run=current_dry_run)
                if sent_ok:
                    success_count += 1
                else:
                    failed_names.append(name)
            except StopRequestedException:
                log(f"\n⚠️ 사용자에 의해 전송이 중단되었습니다. ({i}/{count} 처리됨)")
                stopped = True
                break
            
            # 매크로 탐지 방지를 위한 랜덤 대기 (1~3초, 중단 체크 포함)
            # 패스트 모드는 이 대기를 제거한다(⚠️스팸/매크로 탐지로 계정 제한 위험).
            check_stop_requested()
            try:
                if not current_fast_mode:
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

        try:
            dump_timing_summary()  # B(계측): 단계별 소요시간 요약 + CSV 덤프
        except Exception:
            pass

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
