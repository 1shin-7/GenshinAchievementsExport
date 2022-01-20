
import cv2
import numpy as np

img_path=r"C:\Users\30388\OneDrive\图片\Saved Pictures\reg.png"
img=cv2.imdecode(np.fromfile(img_path,dtype=np.uint8),cv2.IMREAD_COLOR)
cv2.namedWindow('img',0)
cv2.resizeWindow('img',640,480)
up=80
down=1080
left=60
right=600
cv2.imshow('img',img[up:down,left:right])
cv2.waitKey()
