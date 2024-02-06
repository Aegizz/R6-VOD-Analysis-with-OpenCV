from PIL import Image
import xlwt
from xlwt import Workbook
import numpy as np
import os
import cv2

if not os.path.exists('images'):
    os.makedirs('images')

if not os.path.exists('./images/scanned'):
    os.makedirs('./images/scanned')
workbook = xlwt.Workbook() 
wb = Workbook()
sheet1 = wb.add_sheet('Operators')
sheet1.write(0, 0, "Blue")
sheet1.write(0, 2, "Orange")

x = 0
y = 0
i = 0
index = 49021

def find_nearest(array, value):
    array = np.asarray(array)
    idx = (np.abs(array - value)).argmin()
    return array[idx]

def imleft1(opname , x):
      array = [44.13654195, 58.03796347, 61.11127515]
      val = find_nearest(array, opname)

      if (((opname - val) ** 2) ** 0.5) < 1:

            if val == 44.13654195:
                  sheet1.write(x, 0, "SLEDGE")
            
            elif val == 58.03796347:
                  sheet1.write(x, 0, "MUTE")
            
            elif val == 61.11127515:
                  sheet1.write(x, 0, "ALBI")
            
            
      
      else:
            sheet1.write(x, 0, "not in database")
            sheet1.write(x, 5, opname)    

def imright1(opname , y):
      array = [65.0014577866633, 56.595609979591, 47.7601732683577, 56.3769003290433, 97.9257778333125]
      val = find_nearest(array, opname)

      if (((opname - val) ** 2) ** 0.5) < 1:

            if val == 65.0014577866633:
                  sheet1.write(y, 2, "PULSE")
            
            elif val == 56.595609979591:
                  sheet1.write(y, 2, "Melusi")
            
            elif val == 47.7601732683577:
                  sheet1.write(y, 2, "FROST")
            
            elif val == 56.3769003290433:
                  sheet1.write(y, 2, "ZOFIA")
            
            elif val == 97.9257778333125:
                  sheet1.write(y, 2, "IANA")
      
      else:
            sheet1.write(y, 2, "not in database")
            sheet1.write(y, 6, opname)

def imleft2(opname , x):
      array = [56.1945838781794, 32.2321255325431, 58.5266377019446]
      val = find_nearest(array, opname)

      if (((opname - val) ** 2) ** 0.5) < 1:

            if val == 56.1945838781794:
                  sheet1.write(x, 0, "ZOFIA")
            
            elif val == 32.2321255325431:
                  sheet1.write(x, 0, "JÄGER")
            
            elif val == 58.5266377019446:
                  sheet1.write(x, 0, "MUTE")
            
      
      else:
            sheet1.write(x, 0, "not in database")
            sheet1.write(x, 5, opname)
        
def imright2(opname , y):
      array = [66.0065808655088, 75.5860302386605, 71.1978424757383, 44.5297180224083]
      val = find_nearest(array, opname)

      if (((opname - val) ** 2) ** 0.5) < 1:

            if val == 66.0065808655088:
                  sheet1.write(y, 2, "KAID")
            
            elif val == 75.5860302386605:
                  sheet1.write(y, 2, "VALKYRIE")
            
            elif val == 71.1978424757383:
                  sheet1.write(y, 2, "WAMAI")
            
            elif val == 44.5297180224083:
                  sheet1.write(y, 2, "SLEDGE")
      
      else:
            sheet1.write(y, 2, "not in database")
            sheet1.write(y, 6, opname)

def imleft3(opname , x):
      array = [65.4506685789007, 79.9106592989412, 74.149154258236]
      val = find_nearest(array, opname)

      if (((opname - val) ** 2) ** 0.5) < 1:

            if val == 65.4506685789007:
                  sheet1.write(x, 0, "MAVERICK")
            
            elif val == 79.9106592989412:
                  sheet1.write(x, 0, "TACHANKA")
            
            elif val == 74.149154258236:
                  sheet1.write(x, 0, "ARUNI")
            
      
      else:
            sheet1.write(x, 0, "not in database")
            sheet1.write(x, 5, opname)

def imright3(opname , y):
      array = [71.4697821650215, 73.9332333708193, 65.1682285809488]
      val = find_nearest(array, opname)

      if (((opname - val) ** 2) ** 0.5) < 1:

            if val == 71.4697821650215:
                  sheet1.write(y, 2, "WAMAI")
            
            elif val == 73.9332333708193:
                  sheet1.write(y, 2, "ARUNI")
            
            elif val == 65.1682285809488:
                  sheet1.write(y, 2, "MAVERICK")
      
      else:
            sheet1.write(y, 2, "not in database")
            sheet1.write(y, 6, opname)

def imleft4(opname , x):
      array = [44.434091196693, 74.1648036444932, 57.6290968912136]
      val = find_nearest(array, opname)

      if (((opname - val) ** 2) ** 0.5) < 1:

            if val == 44.434091196693:
                  sheet1.write(x, 0, "ACE")
            
            elif val == 74.1648036444932:
                  sheet1.write(x, 0, "ARUNI")
            
            elif val == 57.6290968912136:
                  sheet1.write(x, 0, "MELUSI")
            
      
      else:
            sheet1.write(x, 0, "not in database")
            sheet1.write(x, 5, opname)

def imright4(opname , y):
      array = [32.1599400224916, 78.8369361489441]
      val = find_nearest(array, opname)

      if (((opname - val) ** 2) ** 0.5) < 1:

            if val == 32.1599400224916:
                  sheet1.write(y, 2, "JÄGER")
            
            elif val == 78.8369361489441:
                  sheet1.write(y, 2, "FLORES")
      
      else:
            sheet1.write(y, 2, "not in database")
            sheet1.write(y, 6, opname)

def imleft5(opname , x):
      array = [98.0159024760619, 76.9510271227907]
      val = find_nearest(array, opname)

      if (((opname - val) ** 2) ** 0.5) < 1:

            if val == 98.0159024760619:
                  sheet1.write(x, 0, "IANA")
            
            elif val == 76.9510271227907:
                  sheet1.write(x, 0, "VALYRIE") 
      
      else:
            sheet1.write(x, 0, "not in database")
            sheet1.write(x, 5, opname)

def imright5(opname , y):
      array = [86.12686258, 38.11672185]
      val = find_nearest(array, opname)

      if (((opname - val) ** 2) ** 0.5) < 1:

            if val == 86.12686258:
                  sheet1.write(y, 2, "MUTE")
            
            elif val == 38.11672185:
                  sheet1.write(y, 2, "ACE")
      
      else:
            sheet1.write(y, 2, "not in database")
            sheet1.write(y, 6, opname)
        

# cap = cv2.VideoCapture("VOD.mp4")
# frameTime = 1 # time of each frame in ms, you can add logic to change this value.
# while(cap.isOpened()):
      
#       ret = cap.grab() #grab frame
#       i += 1
#       if i % 60 == 0:
#             ret, frame = cap.retrieve()
#             if not ret: 
#                   break
#             name = './images/' + str(index) + '.jpg'
#             cv2.imwrite(name, frame)
#             index = index + 1


#             if cv2.waitKey(frameTime) & 0xFF == ord('q'):
#                   break

cnt = 0

while cnt < index:
      im = Image.open('./images/' + str(cnt) + ".jpg")

      width, height = im.size
      left = (109/1920) * width
      top = (221/1080) * height
      right = (113/1920) * width
      bottom = (230/1080) *height
      op = im.crop((left, top, right, bottom))
      op_grey = op.convert('LA') 
      width, height = op.size
      total = 0
      for i in range(0, width):
            for j in range(0, height):
                  total += op_grey.getpixel((i,j))[0]
      mean1 = total / (width * height)

      width, height = im.size
      left = (441/1920) * width
      top = (230/1080) * height
      right = (445/1920) * width
      bottom = (240/1080) *height
      op = im.crop((left, top, right, bottom))
      op_grey = op.convert('LA') 
      width, height = op.size
      total = 0
      for i in range(0, width):
            for j in range(0, height):
                  total += op_grey.getpixel((i,j))[0]
      mean2 = total / (width * height)

      width, height = im.size
      left = (305/1920) * width
      top = (250/1080) * height
      right = (309/1920) * width
      bottom = (253/1080) *height
      op = im.crop((left, top, right, bottom))
      op_grey = op.convert('LA') 
      width, height = op.size
      total = 0
      for i in range(0, width):
            for j in range(0, height):
                  total += op_grey.getpixel((i,j))[0]
      mean3 = total / (width * height)

      width, height = im.size
      left = (223/1920) * width
      top = (219/1080) * height
      right = (224/1920) * width
      bottom = (220/1080) *height
      op = im.crop((left, top, right, bottom))
      op_grey = op.convert('LA') 
      width, height = op.size
      total = 0
      for i in range(0, width):
            for j in range(0, height):
                  total += op_grey.getpixel((i,j))[0]
      mean4 = total / (width * height)

      mean = (mean1 + mean2 + mean3 + mean4)/4


      if mean >= 245:
            width, height = im.size
            left = (251/1920) * width
            top = (222/1080) * height
            right = (255/1920) * width
            bottom = (226/1080) *height
            op = im.crop((left, top, right, bottom))
            op_grey = op.convert('LA') 
            width, height = op.size
            total = 0
            for i in range(0, width):
                  for j in range(0, height):
                        total += op_grey.getpixel((i,j))[0]
            mean = total / (width * height)

            if mean < 220:
                  width, height = im.size
                  left = (256/1920) * width
                  top = (244/1080) * height
                  right = (259/1920) * width
                  bottom = (248/1080) *height
                  op = im.crop((left, top, right, bottom))
                  op_grey = op.convert('LA') 
                  width, height = op.size
                  total = 0
                  for i in range(0, width):
                        for j in range(0, height):
                              total += op_grey.getpixel((i,j))[0]
                  mean = total / (width * height)
                  x = x + 1
                  
                  width, height = im.size
                  left = (190/1920) * width
                  top = (300/1080) * height
                  right = (341/1920) * width
                  bottom = (457/1080) * height
                  op = im.crop((left, top, right, bottom))
                  #op.show()
                  op_grey = op.convert('LA') 
                  width, height = op.size
                  total = 0
                  for i in range(0, width):
                        for j in range(0, height):
                              total += op_grey.getpixel((i,j))[0]
                  opname = total / (width * height)
                  imleft1(opname , x)

                  #operator 2
                  x = x + 1
                  width, height = im.size
                  left = (536/1920) * width
                  top = (300/1080) * height
                  right = (687/1920) * width          
                  bottom = (457/1080) *height
                  op = im.crop((left, top, right, bottom))
                  #op.show()
                  op_grey = op.convert('LA') 
                  width, height = op.size
                  total = 0
                  for i in range(0, width):
                        for j in range(0, height):
                              total += op_grey.getpixel((i,j))[0]
                  opname = total / (width * height)
                  imleft2(opname , x)

                  #operator 3
                  x = x + 1
                  width, height = im.size
                  left = (882/1920) * width
                  top = (300/1080) * height
                  right = (1033/1920) * width
                  bottom = (457/1080) *height
                  op = im.crop((left, top, right, bottom))
                  #op.show()
                  op_grey = op.convert('LA') 
                  width, height = op.size
                  total = 0
                  for i in range(0, width):
                        for j in range(0, height):
                              total += op_grey.getpixel((i,j))[0]
                  opname = total / (width * height)
                  imleft3(opname , x)

                  #operator 4
                  x = x + 1
                  width, height = im.size
                  left = (1228/1920) * width
                  top = (300/1080) * height
                  right = (1379/1920) * width
                  bottom = (457/1080) *height
                  op = im.crop((left, top, right, bottom))
                  #op.show()
                  op_grey = op.convert('LA') 
                  width, height = op.size
                  total = 0
                  for i in range(0, width):
                        for j in range(0, height):
                              total += op_grey.getpixel((i,j))[0]
                  opname = total / (width * height)
                  imleft4(opname , x)

                  #operator 5
                  x = x + 1
                  width, height = im.size
                  left = (1574/1920) * width
                  top = (300/1080) * height
                  right = (1725/1920) * width
                  bottom = (457/1080) * height
                  op = im.crop((left, top, right, bottom))
                  #op.show()
                  op_grey = op.convert('LA') 
                  width, height = op.size
                  total = 0
                  for i in range(0, width):
                        for j in range(0, height):
                              total += op_grey.getpixel((i,j))[0]
                  opname = total / (width * height)
                  imleft5(opname , x)

                  #OPERATORS RIGHT SIDE
                  #operator 6
                  y = y + 1
                  width, height = im.size
                  left = (190/1920) * width
                  top = (673/1080) * height
                  right = (341/1920) * width
                  bottom = (832/1080) * height
                  op = im.crop((left, top, right, bottom))
                  #op.show()
                  op_grey = op.convert('LA') 
                  width, height = op.size
                  total = 0
                  for i in range(0, width):
                        for j in range(0, height):
                              total += op_grey.getpixel((i,j))[0]
                  opname = total / (width * height)
                  imright1(opname , y)

                  #operator 7
                  y = y + 1
                  width, height = im.size
                  left = (536/1920) * width
                  top = (673/1080) * height
                  right = (687/1920) * width          
                  bottom = (832/1080) *height
                  op = im.crop((left, top, right, bottom))
                  #op.show()
                  op_grey = op.convert('LA') 
                  width, height = op.size
                  total = 0
                  for i in range(0, width):
                        for j in range(0, height):
                              total += op_grey.getpixel((i,j))[0]
                  opname = total / (width * height)
                  imright2(opname , y)

                  #operator 8
                  y = y + 1
                  width, height = im.size
                  left = (882/1920) * width
                  top = (673/1080) * height
                  right = (1033/1920) * width
                  bottom = (832/1080) *height
                  op = im.crop((left, top, right, bottom))
                  #op.show()
                  op_grey = op.convert('LA') 
                  width, height = op.size
                  total = 0
                  for i in range(0, width):
                        for j in range(0, height):
                              total += op_grey.getpixel((i,j))[0]
                  opname = total / (width * height)
                  imright3(opname , y)

                  #operator 9
                  y = y + 1
                  width, height = im.size
                  left = (1228/1920) * width
                  top = (673/1080) * height
                  right = (1379/1920) * width
                  bottom = (832/1080) *height
                  op = im.crop((left, top, right, bottom))
                  #op.show()
                  op_grey = op.convert('LA') 
                  width, height = op.size
                  total = 0
                  for i in range(0, width):
                        for j in range(0, height):
                              total += op_grey.getpixel((i,j))[0]
                  opname = total / (width * height)
                  imright4(opname , y)

                  #operator 10
                  y = y + 1
                  width, height = im.size
                  left = (1574/1920) * width
                  top = (800/1080) * height
                  right = (1725/1920) * width
                  bottom = (832/1080) * height
                  op = im.crop((left, top, right, bottom))
                  #op.show()
                  op_grey = op.convert('LA') 
                  width, height = op.size
                  total = 0
                  for i in range(0, width):
                        for j in range(0, height):
                              total += op_grey.getpixel((i,j))[0]
                  opname = total / (width * height)
                  imright5(opname , y)

                  im = Image.open('./images/' + str(cnt) + ".jpg")
                  im.save("./images/scanned/" + str(cnt) + ".jpg")
                  cnt = cnt + 90
            else:
                  cnt = cnt + 1
      
      else:
            cnt = cnt + 1

cmt = 0
# while cmt < cnt:
#       os.remove('./images/' + str(cmt) + ".jpg")
#       cmt += 1

wb.save("./images/scanned/success.xls")