import win32com.client
from cvzone.HandTrackingModule import HandDetector
import cv2
import speech_recognition as sr
import threading

# === PowerPoint Setup ===
Application = win32com.client.Dispatch("PowerPoint.Application")
Presentation = Application.Presentations.Open(r"F:\PPTNavigator\Internship PPT.pptx")
print("Opened:", Presentation.Name)
Presentation.SlideShowSettings.Run()

# === Parameters ===
width, height = 900, 720
gestureThreshold = 300
maxZoomFactor = 2
minZoomFactor = 1
zoomFactor = 1

# === Camera Setup ===
cap = cv2.VideoCapture(0)
cap.set(3, width)
cap.set(4, height)

# === Hand Detector ===
detectorHand = HandDetector(detectionCon=0.8, maxHands=1)

# === State Variables ===
delay = 30
buttonPressed = False
counter = 0
imgNumber = 20
annotations = [[]]
annotationNumber = -1
annotationStart = False

# === Voice Command Function ===
def listen_to_voice():
    recognizer = sr.Recognizer()
    mic = sr.Microphone()

    with mic as source:
        recognizer.adjust_for_ambient_noise(source)

    while True:
        with mic as source:
            print("🎤 Listening for command...")
            audio = recognizer.listen(source, phrase_time_limit=4)

        try:
            command = recognizer.recognize_google(audio).lower()
            print(f"🗣️ Voice Command: {command}")

            if "next slide" in command:
                Presentation.SlideShowWindow.View.Next()
                print("➡️ Moved to Next Slide")

            elif "previous slide" in command:
                Presentation.SlideShowWindow.View.Previous()
                print("⬅️ Moved to Previous Slide")

            elif "zoom in" in command:
                global zoomFactor
                zoomFactor = min(maxZoomFactor, zoomFactor + 0.1)
                print(f"🔍 Zoomed In: {zoomFactor:.1f}")

            elif "zoom out" in command:
                zoomFactor = max(minZoomFactor, zoomFactor - 0.1)
                print(f"🔎 Zoomed Out: {zoomFactor:.1f}")

        except sr.UnknownValueError:
            print("❗ Could not understand audio")
        except sr.RequestError:
            print("❗ Could not request results from Google")

# === Start Voice Thread ===
voice_thread = threading.Thread(target=listen_to_voice, daemon=True)
voice_thread.start()

# === Main Loop ===
while True:
    success, img = cap.read()
    imgCurrent = img.copy()

    # Hand Detection
    hands, img = detectorHand.findHands(img)

    if hands and not buttonPressed:
        hand = hands[0]
        cx, cy = hand["center"]
        lmList = hand["lmList"]
        fingers = detectorHand.fingersUp(hand)

        # Gesture Zone
        if cy <= gestureThreshold:
            if fingers == [1, 1, 1, 1, 1]:
                print("Gesture: ➡️ Next Slide")
                buttonPressed = True
                Presentation.SlideShowWindow.View.Next()
                annotations = [[]]
                annotationNumber = -1
                annotationStart = False

            elif fingers == [1, 0, 0, 0, 0]:
                print("Gesture: ⬅️ Previous Slide")
                buttonPressed = True
                Presentation.SlideShowWindow.View.Previous()
                annotations = [[]]
                annotationNumber = -1
                annotationStart = False

    else:
        annotationStart = False

    # Cooldown Timer
    if buttonPressed:
        counter += 1
        if counter > delay:
            counter = 0
            buttonPressed = False

    # Optional Annotation Drawing Preview
    for i, annotation in enumerate(annotations):
        for j in range(len(annotation)):
            if j != 0:
                cv2.line(imgCurrent, annotation[j - 1], annotation[j], (0, 0, 200), 12)

    # Display
    cv2.imshow("Hand & Voice Controlled PPT", img)

    # Exit
    key = cv2.waitKey(1)
    if key == ord('q'):
        break

# Cleanup
cap.release()
cv2.destroyAllWindows()
