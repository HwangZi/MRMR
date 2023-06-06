from PIL import ImageGrab
import time

# 캡처할 화면의 좌표
screen_coordinates = (495, 29.5, 1425, 29.0+930)
# 첫번째 모니터 :  (0, 0, 1920, 1080)


# 화면 캡처를 저장할 파일명
file_name = "simulator.png"

while True:
    #파일 저장 경로 설정 (-> vendor/images/simulator.png로)
    screenshot_file = f"C:/Users/황지혁/Desktop/MRMR/static/vendors/images/{file_name}"

    # 화면 캡처
    screenshot = ImageGrab.grab(bbox=screen_coordinates)

    # 화면 캡처 저장
    screenshot.save(screenshot_file)

    # 캡처 완료 메시지 출력
    print(f"화면 캡처 저장 완료: {screenshot_file}")

    # 5초 대기
    time.sleep(5)