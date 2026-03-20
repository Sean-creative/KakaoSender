"""
카카오톡 자동 메시지 전송기 (macOS) - 웹 버전
- 브라우저 기반 인터페이스 (tkinter 사용 안함)
- Flask 웹 서버 사용
"""

import os
import sys
import subprocess
import threading
import time
import random
import webbrowser
from queue import Queue
from datetime import datetime
from typing import Optional, List

import pandas as pd
import pyperclip
from flask import Flask, render_template_string, request, jsonify, Response
import Quartz
import Vision

# ============================================================
# 설정
# ============================================================
VERSION = "1.0.0"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# 선택 가능한 필터 옵션
AVAILABLE_REGISTER_TYPES = ['이월', '재등록', '신규']
AVAILABLE_AGE_GROUPS = ['10대', '20대', '30대', '40대', '50대', '60대 이상']

# 기본 선택값
DEFAULT_REGISTER_TYPES = ['이월', '재등록', '신규']
DEFAULT_AGE_GROUPS = ['20대', '30대']

# 기본 메시지 템플릿
DEFAULT_MESSAGE_TEMPLATE = "{name}님!\n요청하신 리포트입니다.\n감사합니다."

# Flask 앱
app = Flask(__name__)
log_queue = Queue()
is_running = False
stop_requested = False
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
    delay 0.8
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
        
        # 창 확인
        window_id = get_kakaotalk_window_id()
        if window_id:
            return window_id
        
        time.sleep(1)
    
    return None


def search_friend(name: str):
    """친구 검색 (메뉴 클릭 방식 붙여넣기 사용)"""
    pyperclip.copy(name)
    script = '''
    -- 카카오톡 확실히 활성화
    tell application "KakaoTalk" to activate
    delay 0.3
    
    tell application "System Events"
        tell process "KakaoTalk"
            set frontmost to true
        end tell
        delay 0.3
        
        -- 1. Esc로 혹시 남아있는 채팅창/검색창 닫기
        key code 53
        delay 0.3
        
        -- 2. 친구 목록으로 이동 (Cmd+1) - 채팅창이 아닌 친구 목록에서 검색 보장
        keystroke "1" using command down
        delay 0.5
        
        -- 3. 검색창 열기 (Cmd+F)
        key code 3 using command down
        delay 0.5
        
        -- 4. 기존 검색어 전체 선택 (Cmd+A)
        key code 0 using command down
        delay 0.2
        
        -- 5. 삭제 (Backspace)
        key code 51
        delay 0.3
        
        -- 6. 메뉴 클릭으로 붙여넣기 (Tell Process 필수!)
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
        delay 1.0
        
        -- 7. 아래 화살표 (검색 결과로 이동)
        key code 125
        delay 0.2
        key code 125
        delay 0.2
    end tell
    '''
    run_applescript(script)


def verify_friend_by_ocr(name: str, window_id: int) -> bool:
    """OCR로 친구 검증 - 이름이 완전히 일치하는 텍스트가 최소 1개 이상 있어야 찾은 것으로 인식"""
    texts = capture_and_read(window_id)
    
    # 이름이 완전히 일치하는 텍스트 카운트
    name_count = 0
    for t in texts:
        if name.strip() == t.strip():  # 완전 일치 (앞뒤 공백 제거 후 비교)
            name_count += 1
    
    # 최소 1번 이상 완전 일치해야 친구를 찾은 것으로 인식
    return name_count >= 1


def send_message_to_friend(message: str):
    """채팅방에서 메시지 전송 (메뉴 클릭 방식)"""
    pyperclip.copy(message)
    script = '''
    tell application "KakaoTalk" to activate
    delay 0.3
    tell application "System Events"
        -- 1. 채팅방 열기 (Enter) - 이미 검색 결과에서 화살표로 선택된 상태
        key code 36
        delay 1.0
        
        -- 2. 메시지 붙여넣기 (Menu Click - Robust)
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
        
        -- 3. 전송 (Enter)
        key code 36
        delay 0.5
        
        -- 4. 채팅방 닫기 (Esc)
        key code 53
        delay 0.5
        
        -- 5. (안전장치) 검색창 닫기 (Esc) - 혹시 검색창이 남아있다면
        key code 53
        delay 0.3
    end tell
    '''
    run_applescript(script)

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
        .log-area {
            background: #1e1e1e;
            border-radius: 12px;
            padding: 20px;
            margin-top: 25px;
            max-height: 300px;
            overflow-y: auto;
            font-family: 'Menlo', 'Monaco', monospace;
            font-size: 13px;
        }
        .log-area:empty::before {
            content: "로그가 여기에 표시됩니다...";
            color: #666;
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
            <button class="btn btn-stop" id="stopBtn" style="display:none;" onclick="stopSending()">
                ⏹ 전송 중단
            </button>
            
            <div class="log-area" id="logArea"></div>
        </div>
        <div class="footer">
            카카오톡 자동 전송기 v{{ version }} | 전송 중 마우스/키보드 조작 금지
        </div>
    </div>

    <script>
        let selectedFile = null;
        let eventSource = null;
        
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
            fetch('/stop', { method: 'POST' });
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
                    
                    addLog(data.message, logType);
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
            document.getElementById('stopBtn').style.display = 'none';
            document.getElementById('statusBadge').className = 'status-badge status-idle';
            document.getElementById('statusBadge').textContent = '대기 중';
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
        
        # 임시 파일로 저장
        temp_path = os.path.join(SCRIPT_DIR, 'temp_upload.xlsx')
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
    thread = threading.Thread(target=run_sending_logic, daemon=True)
    thread.start()
    
    return jsonify({'success': True})


@app.route('/stop', methods=['POST'])
def stop_sending():
    global stop_requested
    stop_requested = True
    return jsonify({'success': True})


@app.route('/logs')
def stream_logs():
    def generate():
        while True:
            if not log_queue.empty():
                item = log_queue.get()
                yield f"data: {item}\n\n"
            else:
                time.sleep(0.1)
    
    return Response(generate(), mimetype='text/event-stream')


# ============================================================
# 메시지 전송 로직
# ============================================================
def log(msg):
    """로그 큐에 메시지 추가"""
    import json
    log_queue.put(json.dumps({'type': 'log', 'message': msg}))


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
            log(f"   -> ❌ 카카오톡 창을 찾을 수 없습니다.")
            return False

        # 2. 친구 검색
        log(f"   -> 📋 '{name}' 검색 중...")
        check_stop_requested()
        search_friend(name)
        safe_sleep((1.0, 2.0))  # 검색 결과 로딩 대기 (랜덤, 중단 체크 포함)
        
        # 3. OCR 검증
        check_stop_requested()
        window_id = ensure_kakaotalk_ready()
        check_stop_requested()
        if not window_id:
            log(f"   -> ❌ 카카오톡 창을 찾을 수 없습니다.")
            return False
        
        log(f"   -> 🔍 OCR 검증 중...")
        check_stop_requested()
        if not verify_friend_by_ocr(name, window_id):
            log(f"   -> ❌ '{name}' 친구를 찾을 수 없습니다. (OCR 검증 실패)")
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
                run_applescript(SCRIPT_RESET_SEARCH)
                # 중단 요청이 없을 때만 대기
                if not stop_requested:
                    time.sleep(random.uniform(0.2, 0.5))
            except:
                pass


def run_sending_logic():
    """메인 전송 로직"""
    global is_running, stop_requested
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

        # 동명이인 체크 (앞뒤 공백 제거 후 비교)
        normalized_names = target_df['이름'].str.strip()
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
        
        log("✅ 카카오톡 준비 완료!")
        
        # 친구 목록으로 이동
        run_applescript(SCRIPT_RESET_SEARCH)
        time.sleep(1)
        
        success_count = 0
        failed_names = []
        stopped = False
        
        for i, (_, row) in enumerate(target_df.iterrows()):
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
