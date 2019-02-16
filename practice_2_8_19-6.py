import cv2
import numpy as np

def emptyFunction():
    pass

def main():
    #image = np.zeros((512,512,3),np.uint8)
    image1 = cv2.imread('ella.jpg',cv2.IMREAD_GRAYSCALE)

    windowName = "Ella?"

    cv2.namedWindow(windowName,cv2.WINDOW_GUI_NORMAL)

    cv2.createTrackbar('MinVal',windowName,0,255,emptyFunction)
    cv2.createTrackbar('MaxVal',windowName,0,255,emptyFunction)
    #cv2.createTrackbar('Red',windowName,0,255,emptyFunction)

    while(True):
        #cv2.imshow(windowName,image1)

        if cv2.waitKey(1) == 27:
            break

        minVal1 = cv2.getTrackbarPos('MinVal',windowName)
        maxVal1 = cv2.getTrackbarPos('MaxVal',windowName)
        #red = cv2.getTrackbarPos('Red',windowName)

        #image[:] = [blue,green,red]
        image2=cv2.Canny(image1,minVal1,maxVal1)
        #print(blue,green,red)
        cv2.imshow(windowName,image2)

    cv2.destroyAllWindows()

if __name__ == "__main__":
    main()