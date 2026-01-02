
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

def test_focused_input():
    print("⏳ 3초 뒤에 텍스트가 입력됩니다.")
    print("👉 그 전에 메모장, 브라우저 주소창, 카카오톡 입력창 등 원하는 곳을 클릭해서 커서를 두세요!")
    
    for i in range(3, 0, -1):
        print(f"{i}...")
        time.sleep(1)
    
    print("🚀 Typing now...")
    
    script = '''
    tell application "System Events"
        keystroke "Test Success! (Focus Debug)"
        keystroke return
        keystroke "한글 입력 테스트"
    end tell
    '''
    
    res, err = run_applescript(script)
    
    if err:
        print(f"❌ Error: {err}")
    else:
        print("✅ Input sent.")

if __name__ == "__main__":
    test_focused_input()
