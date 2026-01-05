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
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# 타겟 멤버 필터링 조건
TARGET_REGISTER_TYPES = ['이월', '재등록', '신규']
TARGET_AGE_GROUPS = ['20대', '30대']

# 메시지 템플릿
MESSAGE_TEMPLATE = "{name}님!\n요청하신 리포트입니다.\n감사합니다."

# Flask 앱
app = Flask(__name__)
log_queue = Queue()
is_running = False
current_file_path = None


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
    
    -- 1. Esc 한 번만 (검색창/채팅창 닫기)
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
        
        -- 1. 먼저 친구 목록으로 이동 (Cmd+1) - 안전장치
        keystroke "1" using command down
        delay 0.3
        
        -- 2. 검색창 열기 (Cmd+F)
        key code 3 using command down
        delay 0.5
        
        -- 3. 기존 검색어 전체 선택 (Cmd+A)
        key code 0 using command down
        delay 0.2
        
        -- 4. 삭제 (Backspace)
        key code 51
        delay 0.3
        
        -- 5. 메뉴 클릭으로 붙여넣기 (Tell Process 필수!)
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
        
        -- 6. 아래 화살표 (검색 결과로 이동)
        key code 125
        delay 0.2
        key code 125
        delay 0.2
    end tell
    '''
    run_applescript(script)


def verify_friend_by_ocr(name: str, window_id: int) -> bool:
    """OCR로 친구 검증 - 이름이 최소 2번 나와야 찾은 것으로 인식"""
    texts = capture_and_read(window_id)
    
    # 이름이 포함된 텍스트 카운트
    name_count = 0
    for t in texts:
        if name in t:
            name_count += 1
    
    # 최소 2번 이상 나와야 친구를 찾은 것으로 인식
    return name_count >= 2


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
        .filter-info {
            background: #f8f9fa;
            border-radius: 12px;
            padding: 20px;
            margin-bottom: 25px;
        }
        .filter-info h3 {
            color: #333;
            margin-bottom: 15px;
            font-size: 16px;
        }
        .filter-item {
            display: flex;
            align-items: center;
            gap: 10px;
            margin: 8px 0;
            color: #555;
        }
        .filter-item .badge {
            background: #28a745;
            color: white;
            padding: 4px 10px;
            border-radius: 15px;
            font-size: 12px;
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
            <div class="filter-info">
                <h3>📌 타겟 멤버 필터링 조건</h3>
                <div class="filter-item">
                    <span class="badge">등록형태</span>
                    <span>{{ register_types }}</span>
                </div>
                <div class="filter-item">
                    <span class="badge">연령</span>
                    <span>{{ age_groups }}</span>
                </div>
            </div>
            
            <div class="upload-area" id="uploadArea" onclick="document.getElementById('fileInput').click()">
                <div class="upload-icon">📁</div>
                <div>엑셀 파일을 선택하세요 (.xlsx)</div>
                <div class="file-name" id="fileName"></div>
                <input type="file" id="fileInput" accept=".xlsx" onchange="handleFileSelect(this)">
            </div>
            
            <span class="status-badge status-idle" id="statusBadge">대기 중</span>
            
            <button class="btn btn-start" id="startBtn" disabled onclick="startSending()">
                🚀 카카오톡 전송 시작
            </button>
            
            <div class="log-area" id="logArea"></div>
        </div>
        <div class="footer">
            카카오톡 자동 전송기 v1.0 | 전송 중 마우스/키보드 조작 금지
        </div>
    </div>

    <script>
        let selectedFile = null;
        let eventSource = null;
        
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
            
            const formData = new FormData();
            formData.append('file', selectedFile);
            
            document.getElementById('startBtn').disabled = true;
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
                    fetch('/start', { method: 'POST' });
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
                    else if (data.message.includes('⚠️')) logType = 'warning';
                    else if (data.message.includes('🚀') || data.message.includes('📋')) logType = 'info';
                    
                    addLog(data.message, logType);
                } else if (data.type === 'complete') {
                    eventSource.close();
                    resetUI();
                    if (data.failed_names && data.failed_names.length > 0) {
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
            document.getElementById('startBtn').disabled = false;
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
        register_types=', '.join(TARGET_REGISTER_TYPES),
        age_groups=', '.join(TARGET_AGE_GROUPS)
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
    global is_running
    
    if is_running:
        return jsonify({'success': False, 'error': '이미 실행 중입니다'})
    
    is_running = True
    thread = threading.Thread(target=run_sending_logic, daemon=True)
    thread.start()
    
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


def send_message(name: str, message: str) -> bool:
    """카카오톡 메시지 전송 (OCR 검증 포함)"""
    try:
        # 1. 카카오톡 활성화 및 준비 확인
        window_id = ensure_kakaotalk_ready()
        if not window_id:
            log(f"   -> ❌ 카카오톡 창을 찾을 수 없습니다.")
            return False

        # 2. 친구 검색
        log(f"   -> 📋 '{name}' 검색 중...")
        search_friend(name)
        time.sleep(1.5)  # 검색 결과 로딩 대기
        
        # 3. OCR 검증
        window_id = ensure_kakaotalk_ready()
        if not window_id:
            log(f"   -> ❌ 카카오톡 창을 찾을 수 없습니다.")
            return False
        
        log(f"   -> 🔍 OCR 검증 중...")
        if not verify_friend_by_ocr(name, window_id):
            log(f"   -> ❌ '{name}' 친구를 찾을 수 없습니다. (OCR 검증 실패)")
            return False
        
        log(f"   -> ✅ '{name}' 친구 확인됨!")

        # 4. 메시지 전송
        log(f"   -> 📤 메시지 전송 중...")
        send_message_to_friend(message)
        time.sleep(0.5)
        log(f"   -> ✅ 전송 완료!")
        return True
        
    except Exception as e:
        log(f"   -> ❌ 오류 발생: {e}")
        return False
    finally:
        # 성공/실패 관계없이 다음 검색을 위해 검색창 초기화
        try:
            run_applescript(SCRIPT_RESET_SEARCH)
            time.sleep(0.3)
        except:
            pass


def run_sending_logic():
    """메인 전송 로직"""
    global is_running
    import json
    
    try:
        df = pd.read_excel(current_file_path)
        log(f"📊 전체 {len(df)}명 로드됨")
        
        # 타겟 멤버 필터링
        target_df = df[
            (df['등록형태'].isin(TARGET_REGISTER_TYPES)) &
            (df['연령'].isin(TARGET_AGE_GROUPS))
        ]
        
        count = len(target_df)
        log(f"✅ 타겟 멤버 {count}명 필터링됨")
        
        if count == 0:
            log("⚠️ 타겟 멤버가 없습니다.")
            log_queue.put(json.dumps({
                'type': 'complete',
                'success': 0,
                'total': 0,
                'failed_names': []
            }))
            return
        
        success_count = 0
        failed_names = []
        
        for i, (_, row) in enumerate(target_df.iterrows()):
            name = row['이름']
            message = MESSAGE_TEMPLATE.format(name=name)
            
            log(f"[{i + 1}/{count}] {name}님 처리 중...")
            
            if send_message(name, message):
                success_count += 1
            else:
                failed_names.append(name)
            
            time.sleep(2)
        
        log(f"\n{'='*40}")
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
            'failed_names': failed_names
        }))
        
    except Exception as e:
        log(f"❌ 에러 발생: {e}")
        log_queue.put(json.dumps({
            'type': 'complete',
            'success': 0,
            'total': 0,
            'failed_names': []
        }))
    finally:
        is_running = False


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
