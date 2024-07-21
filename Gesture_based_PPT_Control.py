# This app will use your built-in webcam to control your slides presentation.
# For a one-handed presentation, use Gesture 1 (thumbs up) to go to the previous slide and Gesture 2 (whole hand pointing up) to go to the next slide.

# Import Libraries
import win32com.client
from cvzone.HandTrackingModule import HandDetector
import cv2

Application = win32com.client.Dispatch("PowerPoint.Application")
Presentation = Application.Presentations.Open("C:\\Users\\Dhairya\\Documents\\B. Tech\\SEM-6\\ML\\Demo_PPT.ppt")

print(Presentation.Name)
Presentation.SlideShowSettings.Run()

# Parameters
width, height = 900, 720  #pixels
gestureThreshold = 300

# Camera Setup
cap = cv2.VideoCapture(0)  #default camara
cap.set(3, width)  # 3 for width
cap.set(4, height) # 4 for height

# Hand Detector
detectorHand = HandDetector(detectionCon=0.8, maxHands=1)

# Variables
delay = 30
buttonPressed = False
counter = 0
Slide_Number = 20

while True:
    # Get image frame
    success, img = cap.read()

    # Find the hand and its landmarks
    if success:
        hands, img = detectorHand.findHands(img)  # with draw
        
    if hands and buttonPressed is False:  # If hand is detected
        hand = hands[0]
        cx, cy = hand["center"]
        lmList = hand["lmList"]  # List of 21 Landmark points
        fingers = detectorHand.fingersUp(hand)  # List of which fingers are up
        if cy <= gestureThreshold:  # If hand is at the height of the face
            if fingers == [1, 1, 1, 1, 1]:
                print("Next")
                buttonPressed = True
                if Slide_Number > 0:
                    Presentation.SlideShowWindow.View.Next()
                    Slide_Number -= 1
            if fingers == [1, 0, 0, 0, 0]:
                print("Previous")
                buttonPressed = True
                if Slide_Number > 0:
                    Presentation.SlideShowWindow.View.Previous()
                    Slide_Number += 1
            
    if buttonPressed:
        counter += 1
        if counter > delay:
            counter = 0
            buttonPressed = False

    cv2.imshow("Image", img)
 
    key = cv2.waitKey(1)
    if key == ord('q'):
        break

# Release resources
cap.release()
cv2.destroyAllWindows()




