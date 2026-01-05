
import subprocess
import time
import sys
import Quartz
import Vision
import pyperclip
import pandas as pd # python3 -m pip install pandas openpyxl

# =========================================================
# 1. Automation Helpers (AppleScript)
# =========================================================
def run_applescript(script):
    proc = subprocess.Popen(
        ['osascript', '-'],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE
    )
    result, err = proc.communicate(input=script.encode('utf-8'))
    return result.decode('utf-8'), err.decode('utf-8')

def maximize_window():
    # 'Zoom' 메뉴 아이템 클릭 (윈도우 메뉴 -> 확대/축소)
    script = '''
    tell application "KakaoTalk" to activate
    delay 0.5
    tell application "System Events"
        tell process "KakaoTalk"
            try
                click menu item "확대/축소" of menu "창" of menu bar 1
            on error
                try
                    click menu item "Zoom" of menu "Window" of menu bar 1
                end try
            end try
        end tell
    end tell
    '''
    run_applescript(script)

def go_to_friend_list():
    # Cmd + 1 (친구 목록)
    script = '''
    tell application "System Events"
        keystroke "1" using command down
    end tell
    '''
    run_applescript(script)

def reset_for_next_search():
    """다음 검색을 위해 검색창 초기화 및 친구 목록으로 복귀"""
    script = '''
    tell application "KakaoTalk" to activate
    delay 0.3
    tell application "System Events"
        -- 1. 여러 번 Esc로 모든 창/팝업/검색창 닫기
        key code 53
        delay 0.2
        key code 53
        delay 0.2
        key code 53
        delay 0.3
        
        -- 2. 친구 목록으로 이동 (Cmd+1)
        keystroke "1" using command down
        delay 0.5
    end tell
    '''
    run_applescript(script)

def search_friend(name):
    pyperclip.copy(name)
    script = f'''
    tell application "KakaoTalk" to activate
    delay 0.5
    
    tell application "System Events"
        -- 1. 검색창 열기 (Cmd+F)
        key code 3 using command down
        delay 0.5
        
        -- 2. 기존 검색어 전체 선택 (Cmd+A)
        key code 0 using command down
        delay 0.2
        
        -- 3. 삭제 (Backspace)
        key code 51
        delay 0.3
        
        -- 4. 메뉴 클릭으로 붙여넣기 (Tell Process 필수!)
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
        
        -- 5. 아래 화살표 (검색 결과로 이동)
        key code 125
        delay 0.2
        key code 125
        delay 0.2
    end tell
    '''
    run_applescript(script)

# =========================================================
# 2. Vision / OCR Helpers
# =========================================================
def get_kakaotalk_window_id():
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
                print(f"Found Window: {owner_name} (ID: {window_id})")
                candidates.append(window_id)
                
    return candidates[0] if candidates else None

def capture_and_read(window_id):
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

# =========================================================
# 3. Message Logic
# =========================================================
TARGET_REGISTER_TYPES = ['이월', '재등록', '신규']
TARGET_AGE_GROUPS = ['20대', '30대']
MESSAGE_TEMPLATE = "{name}님!\n요청하신 리포트입니다.\n감사합니다."

def send_message_to_friend(message):
    pyperclip.copy(message)
    script = f'''
    tell application "KakaoTalk" to activate
    delay 0.3
    tell application "System Events"
        -- 1. 채팅방 열기 (Enter) - 이미 검색 결과에서 화살표로 선택된 상태라고 가정
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

# =========================================================
# Main Flow
# =========================================================
def main():
    print(f"\n======== KakaoTalk Auto Sender (Verified) ========")
    
    # 1. Excel Loading
    excel_path = "test_2.xlsx"
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        print(f"❌ Failed to load Excel file: {e}")
        sys.exit(1)
        
    # 2. Filtering
    target_df = df[
        (df['등록형태'].isin(TARGET_REGISTER_TYPES)) &
        (df['연령'].isin(TARGET_AGE_GROUPS))
    ]
    
    total_count = len(target_df)
    print(f"📂 Loaded '{excel_path}'")
    print(f"🎯 Target Members: {total_count} (Filters: {TARGET_REGISTER_TYPES}, {TARGET_AGE_GROUPS})")
    print(f"==================================================\n")

    if total_count == 0:
        print("⚠️ No targets found.")
        return

    # 3. Main Loop
    success_count = 0
    fail_count = 0
    
    print("[Step 1] Initializing KakaoTalk...")
    run_applescript('tell application "KakaoTalk" to activate')
    maximize_window()
    time.sleep(1)
    
    # 친구 목록으로 한 번만 이동해두면 계속 거기서 검색 가능? 
    # -> 검색창 닫기(Esc) 하면 다시 친구 목록 탭이 보여야 함.
    print("[Step 2] Going to Friend List...")
    go_to_friend_list()
    time.sleep(1)
    
    for idx, (_, row) in enumerate(target_df.iterrows()):
        name = row['이름']
        msg = MESSAGE_TEMPLATE.format(name=name)
        
        print(f"\n[{idx+1}/{total_count}] Processing: {name} ...")
        
        # A. 검색
        search_friend(name)
        time.sleep(1.5) # 검색 결과 로딩 대기
        
        # B. 검증 (OCR)
        window_id = get_kakaotalk_window_id()
        if not window_id:
            print("   ❌ KakaoTalk Window Not Found")
            fail_count += 1
            # 다음 검색을 위해 초기화
            reset_for_next_search()
            time.sleep(1)
            continue
            
        texts = capture_and_read(window_id)
        
        # OCR 결과 콘솔 출력
        print(f"   📷 OCR 결과 (총 {len(texts)}개):")
        for i, t in enumerate(texts):
            print(f"      [{i+1}] {t}")
        
        # B-1. 필터링 로직
        filtered_texts = []
        for t in texts:
            t_clean = t.strip()
            if t_clean in ["채팅", "친구", "...", "..", "•", "2", "8", "Q"]: continue
            if t_clean.startswith("Q") and name in t_clean: continue
            if t_clean == name: continue
            filtered_texts.append(t_clean)
        
        # 필터링 후 결과도 출력
        print(f"   🔍 필터링 후 ({len(filtered_texts)}개): {filtered_texts}")
            
        # B-2. 판단
        found_by_name = any(name in ft for ft in filtered_texts)
        meaningful_line_count = len([ft for ft in filtered_texts if len(ft) >= 2])
        found_by_density = meaningful_line_count >= 2
        
        is_found = found_by_name or found_by_density
        
        if is_found:
            print(f"   ✅ Verified! (Name={found_by_name}, Density={found_by_density}, Lines={meaningful_line_count})")
            
            # C. 전송
            print("   📤 Sending Message...")
            send_message_to_friend(msg)
            print("   ✅ Sent.")
            success_count += 1
            
        else:
            print(f"   ❌ Not Found (Density Lines: {meaningful_line_count}). Skipping.")
            fail_count += 1
            
        # 성공/실패 관계없이 다음 검색을 위해 검색창 초기화
        reset_for_next_search()
            
        time.sleep(1) # Interval
        
    print(f"\n{'='*40}")
    print(f"🎉 Completed! Success: {success_count}, Failed: {fail_count}")
    print(f"{'='*40}")

if __name__ == "__main__":
    main()
