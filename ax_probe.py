"""
카카오톡 macOS 접근성(AX) 탐침 스크립트 — 읽기 전용

목적:
  - OCR 대체 가능성 검증용.
  - 카카오톡 창의 접근성(AXUIElement) 트리를 덤프해서,
    검색 결과의 친구 이름이 '텍스트'(AXStaticText 등)로 노출되는지 확인한다.

안전성:
  - 클릭/키입력/메시지 전송 등 부수효과(side effect)가 전혀 없다.
  - 오직 읽기(Copy attribute)만 수행한다.

사용법:
  1) 카카오톡을 실행하고, 확인하고 싶은 화면(예: 친구 검색 결과)을 띄워 둔다.
       - 검색 결과를 보려면 카카오톡에서 직접 이름을 검색해 결과가 보이는 상태로 둔다.
  2) 터미널에서:  python3 ax_probe.py
       - 이름이 들어간 노드만 보고 싶으면:  python3 ax_probe.py --text-only
       - 트리 깊이 제한 변경:                python3 ax_probe.py --max-depth 25
  3) 처음 실행 시 '시스템 설정 > 개인정보 보호 및 보안 > 접근성'에서
     터미널(또는 실행 주체)에 권한을 부여해야 한다.
"""

import sys
import argparse

try:
    from AppKit import NSWorkspace
    from ApplicationServices import (
        AXUIElementCreateApplication,
        AXUIElementCopyAttributeValue,
        AXUIElementCopyAttributeNames,
        AXIsProcessTrusted,
    )
    import Quartz
except Exception as exc:  # pragma: no cover - 환경 의존
    print(f"[!] 필요한 PyObjC 프레임워크를 불러오지 못했습니다: {exc}")
    print("    pip install pyobjc 가 되어 있는지 확인하세요.")
    sys.exit(1)


KAKAO_BUNDLE_IDS = {"com.kakao.KakaoTalkMac", "com.kakao.KakaoTalk"}
KAKAO_NAME_HINTS = ("KakaoTalk", "카카오톡")

# 관심 있는 텍스트 계열 role (이름이 여기에 담길 가능성이 높음)
TEXT_ROLES = {"AXStaticText", "AXTextField", "AXTextArea"}
# 구조적으로 의미 있는 컨테이너 role (검색 결과 리스트 추적용)
STRUCTURE_ROLES = {"AXScrollArea", "AXTable", "AXRow", "AXCell", "AXOutline", "AXList", "AXGroup"}

ATTRS_OF_INTEREST = ["AXRole", "AXSubrole", "AXIdentifier", "AXTitle", "AXValue", "AXDescription"]


def _copy_attr(element, attr):
    """AX 속성 한 개를 안전하게 읽어 반환. 실패 시 None."""
    try:
        err, value = AXUIElementCopyAttributeValue(element, attr, None)
        if err != 0:
            return None
        return value
    except Exception:
        return None


def _attr_names(element):
    try:
        err, names = AXUIElementCopyAttributeNames(element, None)
        if err != 0 or names is None:
            return []
        return list(names)
    except Exception:
        return []


def find_kakao_pid():
    """실행 중인 카카오톡의 PID를 찾는다. 없으면 None."""
    workspace = NSWorkspace.sharedWorkspace()
    for app in workspace.runningApplications():
        bundle_id = app.bundleIdentifier() or ""
        name = app.localizedName() or ""
        if bundle_id in KAKAO_BUNDLE_IDS or any(h in name for h in KAKAO_NAME_HINTS):
            return app.processIdentifier(), (bundle_id or name)
    return None, None


def describe(element):
    """한 노드를 사람이 읽기 좋은 한 줄 설명으로 만든다."""
    role = _copy_attr(element, "AXRole") or "?"
    parts = [str(role)]
    for attr in ("AXSubrole", "AXIdentifier", "AXTitle", "AXValue", "AXDescription"):
        val = _copy_attr(element, attr)
        if val is None:
            continue
        text = str(val).replace("\n", " ").strip()
        if not text:
            continue
        if len(text) > 60:
            text = text[:60] + "…"
        parts.append(f"{attr}='{text}'")
    return "  ".join(parts)


def has_text(element):
    role = _copy_attr(element, "AXRole")
    if role in TEXT_ROLES:
        return True
    for attr in ("AXValue", "AXTitle", "AXDescription"):
        val = _copy_attr(element, attr)
        if val and str(val).strip():
            return True
    return False


class Stats:
    def __init__(self):
        self.nodes = 0
        self.text_nodes = 0
        self.text_samples = []


def walk(element, depth, max_depth, text_only, stats):
    stats.nodes += 1

    node_has_text = has_text(element)
    if node_has_text:
        stats.text_nodes += 1
        role = _copy_attr(element, "AXRole")
        if role in TEXT_ROLES:
            for attr in ("AXValue", "AXTitle", "AXDescription"):
                val = _copy_attr(element, attr)
                if val and str(val).strip():
                    sample = str(val).replace("\n", " ").strip()
                    if sample and sample not in stats.text_samples:
                        stats.text_samples.append(sample)
                    break

    if (not text_only) or node_has_text:
        print(f"{'  ' * depth}- {describe(element)}")

    if depth >= max_depth:
        return

    children = _copy_attr(element, "AXChildren") or []
    for child in children:
        walk(child, depth + 1, max_depth, text_only, stats)


def main():
    parser = argparse.ArgumentParser(description="카카오톡 AX 트리 읽기 전용 탐침")
    parser.add_argument("--max-depth", type=int, default=20, help="트리 최대 깊이 (기본 20)")
    parser.add_argument("--text-only", action="store_true", help="텍스트가 있는 노드만 출력")
    args = parser.parse_args()

    print("=" * 60)
    print(" 카카오톡 접근성(AX) 탐침 — 읽기 전용 (전송/클릭 없음)")
    print("=" * 60)

    if not AXIsProcessTrusted():
        print("[!] 접근성 권한이 없습니다.")
        print("    시스템 설정 > 개인정보 보호 및 보안 > 접근성 에서")
        print("    이 스크립트를 실행하는 앱(터미널/IDE)에 권한을 부여한 뒤 다시 실행하세요.")
        return

    pid, ident = find_kakao_pid()
    if not pid:
        print("[!] 실행 중인 카카오톡을 찾지 못했습니다. 카카오톡을 먼저 실행하세요.")
        return

    print(f"[+] 카카오톡 발견: pid={pid}, ident={ident}")
    app_element = AXUIElementCreateApplication(pid)
    if app_element is None:
        print("[!] AXUIElement 생성 실패.")
        return

    windows = _copy_attr(app_element, "AXWindows") or []
    print(f"[+] 최상위 창 개수(AXWindows): {len(windows)}")
    print(f"[+] 앱 레벨 속성: {_attr_names(app_element)}")

    # AXWindows가 비어 있어도 AXMainWindow/AXFocusedWindow로 창이 잡힐 수 있다.
    roots = list(windows)
    seen_ids = set(id(w) for w in roots)
    for attr in ("AXMainWindow", "AXFocusedWindow"):
        w = _copy_attr(app_element, attr)
        if w is not None and id(w) not in seen_ids:
            print(f"[+] {attr} 로 창 발견 → 순회 대상에 추가")
            roots.append(w)
            seen_ids.add(id(w))
    print("-" * 60)

    stats = Stats()
    if not roots:
        print("[!] 창을 전혀 잡지 못했습니다. 앱 요소(메뉴바 포함)부터 순회합니다.")
        walk(app_element, 0, args.max_depth, args.text_only, stats)
    else:
        for idx, win in enumerate(roots):
            title = _copy_attr(win, "AXTitle") or "(제목 없음)"
            role = _copy_attr(win, "AXRole") or "?"
            print(f"\n##### 창 [{idx}] role={role} title='{title}' #####")
            walk(win, 0, args.max_depth, args.text_only, stats)

    print("\n" + "=" * 60)
    print(f" 요약: 전체 노드 {stats.nodes}개, 텍스트 보유 노드 {stats.text_nodes}개")
    print("=" * 60)
    if stats.text_samples:
        print(" 읽힌 텍스트 샘플 (최대 40개):")
        for s in stats.text_samples[:40]:
            print(f"   • {s}")
    else:
        print(" [!] 텍스트를 가진 노드를 찾지 못했습니다.")
        print("     → 이 화면은 AX로 텍스트가 노출되지 않을 수 있습니다 (OCR 폴백 필요).")


if __name__ == "__main__":
    main()
