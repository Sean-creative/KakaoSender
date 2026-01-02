
import time
import sys
import Quartz
from AppKit import NSWorkspace
from ApplicationServices import (
    AXUIElementCreateApplication, 
    kAXChildrenAttribute, 
    kAXRoleAttribute, 
    kAXTitleAttribute, 
    kAXValueAttribute,
    AXUIElementCopyAttributeValue,
    AXUIElementCreateApplication, 
    kAXChildrenAttribute, 
    kAXRoleAttribute, 
    kAXTitleAttribute, 
    kAXValueAttribute,
    AXUIElementCopyAttributeValue,
    AXUIElementCopyAttributeNames,
    AXIsProcessTrusted
)

def get_kakaotalk_pid():
    """카카오톡 실행 중인지 확인하고 PID 반환"""
    workspace = NSWorkspace.sharedWorkspace()
    for app in workspace.runningApplications():
        name = app.localizedName() or ""
        bid = app.bundleIdentifier() or ""
        
        # 카카오톡 찾기 (한글 이름 '카카오톡' 또는 영문 'KakaoTalk', Bundle ID 포함)
        if "KakaoTalk" in name or "카카오톡" in name or "com.kakao.KakaoTalk" in bid:
            print(f"Found App: {name} ({bid})")
            return app.processIdentifier()
    return None

def traverse_ax_element(element, depth=0, max_depth=5):
    """재귀적으로 UI 요소 트리 탐색"""
    if depth > max_depth:
        return

    indent = "  " * depth
    
    # 기본 속성
    try:
        _, role = AXUIElementCopyAttributeValue(element, kAXRoleAttribute, None)
    except: return

    try:
        _, title = AXUIElementCopyAttributeValue(element, kAXTitleAttribute, None)
    except: title = ""
        
    try:
        _, value = AXUIElementCopyAttributeValue(element, kAXValueAttribute, None)
    except: value = ""

    # 출력
    print(f"{indent}[{role}] Title: '{title}', Value: '{value}'")

    # 자식 요소 탐색
    try:
        _, children = AXUIElementCopyAttributeValue(element, kAXChildrenAttribute, None)
        if children:
            for child in children:
                # 만약 Row라면 그 자식들은 무조건 출력해본다 (Depth 무시하고 1단계 더)
                if role == "AXRow":
                    print(f"{indent}  >>> Found Row! Inspecting children...")
                    for sub_child in children:
                        traverse_ax_element(sub_child, depth + 1, depth + 2)
                    return # 첫 번째 Row만 보고 종료 (너무 길어지므로)
                
                traverse_ax_element(child, depth + 1, max_depth)
    except:
        pass

def main():
    print("CoreGraphics/Accessibility API Test for KakaoTalk")
    print("=================================================")
    
    if not AXIsProcessTrusted():
        print("⚠️  WARNING: Process is NOT trusted (AXIsProcessTrusted = False).")
        print("⚠️  You must grant 'Accessibility' permission to this terminal/application.")
        print("    Go to System Settings > Privacy & Security > Accessibility and add your terminal.")
    else:
        print("✅ Process is trusted.")

    pid = get_kakaotalk_pid()
    if not pid:
        print("❌ KakaoTalk is not running. Please open KakaoTalk first.")
        return

    print(f"✅ Found KakaoTalk (PID: {pid})")
    print("⏳ Connecting to Accessibility API...")
    
    app_ref = AXUIElementCreateApplication(pid)
    
    print("\n🔍 Dumping UI Hierarchy (Depth: 4)...")
    print("-" * 50)
    
    # 윈도우 목록 혹은 전체 앱 자식 탐색
    traverse_ax_element(app_ref, max_depth=4)
    
    print("-" * 50)
    print("Done.")

if __name__ == "__main__":
    try:
        main()
    except ImportError:
        print("❌ Required libraries not found.")
        print("Please run: pip install pyobjc")
    except Exception as e:
        print(f"❌ An error occurred: {e}")
        print("\nNote: You may need to grant 'Accessibility' permission to your Terminal/Python.")
        print("System Settings > Privacy & Security > Accessibility")
