
import subprocess
import time
import sys

def run_applescript(script):
    proc = subprocess.Popen(
        ['osascript', '-'],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE
    )
    result, err = proc.communicate(input=script.encode('utf-8'))
    return result.decode('utf-8'), err.decode('utf-8')

def test_menu_paste():
    print("📋 '메뉴 클릭(Edit -> Paste)' 방식 테스트")
    print("1. 텍스트를 복사(Cmd+C) 하세요.")
    print("2. 3초 뒤에 [편집] -> [붙여넣기] 메뉴를 클릭합니다.")
    print("3. 메모장이나 카카오톡을 활성화해서 커서를 두세요!")
    
    for i in range(3, 0, -1):
        print(f"{i}...")
        time.sleep(1)
    
    print("🚀 Clicking 'Paste' menu on KakaoTalk...")
    
    script = '''
    tell application "KakaoTalk" to activate
    delay 0.5
    
    tell application "System Events"
        tell process "KakaoTalk"
            set frontmost to true
            try
                -- 1. 한글 메뉴 (편집 -> 붙여넣기)
                click menu item "붙여넣기" of menu "편집" of menu bar 1
            on error
                try
                    -- 2. 혼합 메뉴 (편집 -> Paste) **[스크린샷 기준 유력]**
                    click menu item "Paste" of menu "편집" of menu bar 1
                on error
                    try
                        -- 3. 영문 메뉴 (Edit -> Paste)
                        click menu item "Paste" of menu "Edit" of menu bar 1
                    on error
                        return "Menu item not found"
                    end try
                end try
            end try
        end tell
    end tell
    '''
    
    res, err = run_applescript(script)
    
    if "Menu item not found" in res:
        print("❌ Failed: 'Paste' menu item not found.")
    elif err:
        print(f"❌ Error: {err}")
    else:
        print("✅ Menu click sent.")

if __name__ == "__main__":
    test_menu_paste()
