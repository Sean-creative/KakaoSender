
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
    kAXPlaceholderValueAttribute,
    AXUIElementCopyAttributeValue,
    AXUIElementSetAttributeValue,
    AXIsProcessTrusted
)

def get_kakaotalk_pid():
    workspace = NSWorkspace.sharedWorkspace()
    for app in workspace.runningApplications():
        name = app.localizedName() or ""
        if "KakaoTalk" in name or "카카오톡" in name:
            return app.processIdentifier()
    return None

def find_search_field_and_write(element, depth=0, max_depth=10):
    if depth > max_depth:
        return False

    # Check Role
    try:
        _, role = AXUIElementCopyAttributeValue(element, kAXRoleAttribute, None)
    except: return False
    
    # Check if this is a text field
    if role in ["AXTextField", "AXSearchField"]:
        print(f"🎯 Found Field Candidate! Role: {role}")
        
        # Try writing "Test"
        try:
            er = AXUIElementSetAttributeValue(element, kAXValueAttribute, "김선우")
            if er == 0: # kAXErrorSuccess
                print("✅ Successfully wrote to this field!")
                return True
            else:
                print(f"❌ Write failed (Error Code: {er})")
        except Exception as e:
            print(f"❌ Write exception: {e}")

    # Traverse Children
    try:
        _, children = AXUIElementCopyAttributeValue(element, kAXChildrenAttribute, None)
        if children:
            for child in children:
                if find_search_field_and_write(child, depth + 1, max_depth):
                    return True
    except:
        pass
        
    return False

def main():
    print("🔍 Hunting for KakaoTalk Search Field...")
    pid = get_kakaotalk_pid()
    if not pid:
        print("❌ KakaoTalk not running")
        return
        
    app_ref = AXUIElementCreateApplication(pid)
    found = find_search_field_and_write(app_ref, max_depth=10)
    
    if not found:
        print("❌ Could not find/write to any search field via Accessibility.")

if __name__ == "__main__":
    main()
