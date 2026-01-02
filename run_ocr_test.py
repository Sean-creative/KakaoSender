
import Quartz
import Vision
from Cocoa import NSURL
import sys

def get_kakaotalk_window_id():
    """
    카카오톡의 메인 윈도우 ID를 찾습니다.
    (화면상에 있는 윈도우만 검색)
    """
    options = Quartz.kCGWindowListOptionOnScreenOnly
    window_list = Quartz.CGWindowListCopyWindowInfo(options, Quartz.kCGNullWindowID)
    
    candidates = []
    
    for window in window_list:
        owner_name = window.get('kCGWindowOwnerName', '')
        title = window.get('kCGWindowName', '')
        window_id = window.get('kCGWindowNumber', 0)
        bounds = window.get('kCGWindowBounds', {})
        
        # 카카오톡 프로세스 찾기 (한글/영문)
        if 'KakaoTalk' in owner_name or '카카오톡' in owner_name:
            # 너무 작은 윈도우(알림창, 투명창 등) 제외
            width = bounds.get('Width', 0)
            height = bounds.get('Height', 0)
            
            if width > 200 and height > 300:
                print(f"Found Window: ID={window_id}, Owner={owner_name}, Title='{title}', Size={width}x{height}")
                candidates.append(window_id)

    if not candidates:
        return None
    
    # 여러 개라면 가장 마지막(보통 활성화된) 윈도우 혹은 첫번째 반환
    # (여기서는 단순히 첫 번째 발견된 적절한 크기의 윈도우 사용)
    return candidates[0]

def capture_window(window_id):
    """
    특정 윈도우 ID만 캡처하여 CGImageRef 생성
    """
    image_ref = Quartz.CGWindowListCreateImage(
        Quartz.CGRectNull,
        Quartz.kCGWindowListOptionIncludingWindow,
        window_id,
        Quartz.kCGWindowImageBoundsIgnoreFraming | Quartz.kCGWindowImageNominalResolution
    )
    return image_ref

def recognize_text(cg_image):
    """
    Vision Framework를 사용하여 이미지에서 텍스트 추출
    """
    # 요청 핸들러 생성
    request_handler = Vision.VNImageRequestHandler.alloc().initWithCGImage_options_(cg_image, None)
    
    # 텍스트 인식 요청 생성
    request = Vision.VNRecognizeTextRequest.alloc().init()
    request.setRecognitionLevel_(Vision.VNRequestTextRecognitionLevelAccurate) # 정확도 우선
    request.setUsesLanguageCorrection_(True) # 언어 보정 사용
    request.setRecognitionLanguages_(['ko-KR', 'en-US']) # 한국어/영어
    
    # 실행 (PyObjC에서는 (Bool, Error) 튜플 반환)
    success, error_obj = request_handler.performRequests_error_([request], None)
    
    if success:
        results = request.results()
        if not results:
            print("No text detected.")
            return

        print("\n" + "="*40)
        print("🔍 Detected Text Results:")
        print("="*40)
        
        full_text = []
        for observation in results:
            # 후보군 중 가장 신뢰도 높은 첫 번째 녀석
            top_candidate = observation.topCandidates_(1)[0]
            text = top_candidate.string()
            confidence = top_candidate.confidence()
            
            print(f"[{confidence:.2f}] {text}")
            full_text.append(text)
            
        return full_text
    else:
        print(f"Error during text recognition: {error_obj}")

def main():
    print("🚀 searching for KakaoTalk window...")

    # 권한 체크 (macOS 10.15+)
    if hasattr(Quartz, 'CGPreflightScreenCaptureAccess'):
        has_access = Quartz.CGPreflightScreenCaptureAccess()
        print(f"🔒 Screen Capture Access: {has_access}")
        if not has_access:
            print("⚠️ Requesting Screen Recording permission...")
            # 권한 요청 (시스템 팝업 뜸)
            Quartz.CGRequestScreenCaptureAccess()
            print("❌ Please allow 'Screen Recording' in System Settings > Privacy & Security.")
            return
    
    # 1. 윈도우 찾기
    window_id = get_kakaotalk_window_id()
    if not window_id:
        print("❌ Could not find KakaoTalk window.")
        print("Make sure KakaoTalk is OPEN and VISIBLE on any screen.")
        return

    print(f"✅ Target Window ID: {window_id}")
    
    # 2. 캡처
    print("📸 Capturing window image...")
    cg_image = capture_window(window_id)
    
    if not cg_image:
        print("❌ Failed to capture window.")
        return

    # 3. OCR 수행
    print("👀 Reading text via Vision Framework...")
    recognize_text(cg_image)

if __name__ == "__main__":
    main()
