
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

def test_paste_only():
    print("📋 '붙여넣기(Cmd+V)' 테스트 모드")
    print("1. 지금 원하는 텍스트를 복사(Cmd+C) 하세요.")
    print("2. 3초 뒤에 붙여넣기 명령이 실행됩니다.")
    print("3. 그 전에 메모장이나 카카오톡 입력창에 커서를 두세요!")
    
    for i in range(3, 0, -1):
        print(f"{i}...")
        time.sleep(1)
    
    print("🚀 Pasting now (Cmd+V)...")
    
    # System Events를 통해 현재 활성화된 앱에 Cmd+V 전송
    script = '''
    tell application "System Events"
        keystroke "v" using command down
    end tell
    '''
    
    res, err = run_applescript(script)
    
    if err:
        print(f"❌ Error: {err}")
    else:
        print("✅ Paste command sent.")

if __name__ == "__main__":
    test_paste_only()
