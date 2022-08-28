
import os
import cv2
import pytesseract
from xlwt import Workbook
import datetime

#video Capture
vidcap = cv2.VideoCapture('short test video.mp4')
success,image = vidcap.read()
count = 1
cnt=1
txt1=""
try:
 os.remove("frame.jpg")
except:
 pass

wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1')

#image to text
while success:
 cv2.imwrite("frame.jpg" , image)
 success,image = vidcap.read()
 pytesseract.pytesseract.tesseract_cmd ='C:\\Program Files\\Tesseract-OCR\\tesseract.exe'
 img = cv2.imread("frame.jpg", cv2.COLOR_BGR2GRAY)
 txt2=pytesseract.image_to_string(img)
 if txt1==txt2:
   continue
 else:
   txt1=txt2
   print(txt1+" "+str(count))
   data = txt1.split()
   print(data)
   floats = []
   for elem in data:
       try:
           floats.append(float(elem))
       except ValueError:
           pass
     #Save text to xl
   print(floats)
   for i in range(0,len(floats)):
     sheet1.write(count,i, floats[i])

 if(count==50):
   wb.save('xlwt'+str(cnt)+'.xls')
   count=0
   cnt+=1
   sheet1 = wb.add_sheet('Sheet'+str(cnt))

 count += 1
 try:
   os.remove("frame.jpg")
 except:
   pass
