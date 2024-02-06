from PIL import Image, ImageEnhance, ImageFilter
import xlwt
from xlwt import Workbook
import numpy as np
import os
import cv2
from sys import path
from typing import Text
import pytesseract
import pytesseract as ocr
import pytesseract
import math
import argparse
import difflib

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
ap = argparse.ArgumentParser()

ap.add_argument("-p", "--preprocess", type=str, default="thresh",	help="type of preprocessing to be done")
args = vars(ap.parse_args())

if not os.path.exists('images'):
    os.makedirs('images')

if not os.path.exists('./images/scanned'):
    os.makedirs('./images/scanned')

if not os.path.exists('./images/banned'):
    os.makedirs('./images/banned')

workbook = xlwt.Workbook() 
wb = Workbook()

sheet1 = wb.add_sheet('Operators')
sheet1.write(0, 0, "Blue")
sheet1.write(0, 2, "Orange")

sheet2 = wb.add_sheet('Site')
sheet2.write(0, 0, "Map")
sheet2.write(0, 2, "Site")

sheet3 = wb.add_sheet('Banned')
sheet3.write(0, 0, "BLUE")
sheet3.write(0, 2, "ORANGE")


operator_names = ["RECRUIT","SLEDGE","THATCHER","ASH","THERMITE","TWITCH","MONTAGNE","GLAZ","FUZE","BLITZ","IQ","BUCK","BLACKBEARD" ,"CAPITÃO" ,"HIBANA","JACKAL","YING","ZOFIA","DOKKAEBI","LION","FINKA","MAVERICK","NOMAD","GRIDLOCK","NØKK" ,"AMARU","KALI" ,"IANA","ACE","ZERO","FLORES","OSA","SMOKE","MUTE","CASTLE","PULSE","DOC","ROOK","KAPKAN","TACHANKA","JÄGER","BANDIT","FROST" ,"VALKYRIE","CAVEIRA","ECHO","MIRA","LESION","ELA" ,"VIGIL","ALIBI","MAESTRO","CLASH","KAID","MOZZIE","WARDEN","GOYO","WAMAI","ORYX","MELUSI","ARUNI","THUNDERBIRD"]


p = 1
q = 1
gate = 1
x = 1
y = 1
i = 0
r = 1
label = 0
count = 0
vidcap = cv2.VideoCapture("VOD.mp4")
success,image = vidcap.read()
while success == True:
      vidcap.set(cv2.CAP_PROP_POS_MSEC,(count*1000))    
      success,image = vidcap.read()
      if success == True:
            cv2.imwrite("./images/" + "\\%d.jpg" % count, image)    
            im = Image.open('./images/' + str(count) + ".jpg")

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
                        label = label + 1
                        im.save("./images/scanned/" + str(label) + ".jpg")
                        
                        image = cv2.imread("./images/scanned/" + str(label) + ".jpg")
                        im = Image.open("./images/scanned/" + str(label) + ".jpg")

                        width, height = im.size
                        left = (1189/1920) * width
                        top = (82/1080) * height
                        right = (1190/1920) * width
                        bottom = (83/1080) *height
                        op = im.crop((left, top, right, bottom))
                        op_grey = op.convert('LA') 
                        width, height = op.size
                        total = 0
                        for i in range(0, width):
                              for j in range(0, height):
                                    total += op_grey.getpixel((i,j))[0]
                        men = total / (width * height)
                        print(men)


                        if men > 240:
                              #blue
                              crop_img = image[950:1016, 840:1126]
                              gray = cv2.cvtColor(crop_img, cv2.COLOR_BGR2GRAY)
                              if args["preprocess"] == "thresh":
                                    gray = cv2.threshold(gray, 0, 255,
                                          cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
                              elif args["preprocess"] == "blur":
                                    gray = cv2.medianBlur(gray, 3)
                              filename = "{}.png".format(os.getpid())
                              cv2.imwrite(filename, gray)
                              text = pytesseract.image_to_string(Image.open(filename))
                              sheet2.write (r, 1, text)
                              os.remove(filename)
                              r = r + 5

                        else:
                              #orange
                              cropped_image = image[579:644, 798:1125]
                              
                              gray2 = cv2.cvtColor(cropped_image, cv2.COLOR_BGR2GRAY)
                              if args["preprocess"] == "thresh":
                                    gray2 = cv2.threshold(gray2, 0, 255,
                                          cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
                              elif args["preprocess"] == "blur":
                                    gray2 = cv2.medianBlur(gray2, 3)
                              filename2 = "{}.png".format(os.getpid())
                              cv2.imwrite(filename2, gray2)
                              text2 = pytesseract.image_to_string(Image.open(filename2))
                              sheet2.write (r, 3, text2)
                              os.remove(filename2)
                              r = r + 5
                  
                        #Operators 1-Blue
                        crop_image2 = image[478:45+478, 175:175+270]
                        gray4 = cv2.cvtColor(crop_image2, cv2.COLOR_BGR2GRAY)
                        if args["preprocess"] == "thresh":
                              gray4 = cv2.threshold(gray4, 0, 255,
                                    cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
                        elif args["preprocess"] == "blur":
                              gray4 = cv2.medianBlur(gray4, 3)
                        filename4 = "{}.png".format(os.getpid())
                        thresh2 = cv2.threshold(crop_image2, 150, 255, cv2.THRESH_BINARY_INV)[1]

                        result2 = cv2.GaussianBlur(thresh2, (5,5), 0)
                        result2 = 255 - result2

                        operator = pytesseract.image_to_string(result2, lang='eng',config='--psm 6')
                        opclose = difflib.get_close_matches(operator, operator_names, 1, 0.4)
                        print(opclose)
                        sheet1.write (x, 0, opclose)
                        x=x+1

                        #Operators 2-Blue
                        crop_image3 = image[478:45+478, 515:515+270]
                  
                        cv2.waitKey(0)
                        gray5 = cv2.cvtColor(crop_image3, cv2.COLOR_BGR2GRAY)
                        if args["preprocess"] == "thresh":
                              gray5 = cv2.threshold(gray5, 0, 255,
                                    cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
                        elif args["preprocess"] == "blur":
                              gray5 = cv2.medianBlur(gray5, 3)
                        filename5 = "{}.png".format(os.getpid())
                        thresh3 = cv2.threshold(crop_image3, 150, 255, cv2.THRESH_BINARY_INV)[1]

                        result3 = cv2.GaussianBlur(thresh3, (5,5), 0)
                        result3 = 255 - result3

                        operator = pytesseract.image_to_string(result3, lang='eng',config='--psm 6')
                        opclose = difflib.get_close_matches(operator, operator_names, 1, 0.4)
                        print(opclose)
                        sheet1.write (x, 0, opclose)
                        x=x+1
                        #Operator 3-Blue
                        crop_image4 = image[478:45+478, 515+350:515+340+270]
                        
                        cv2.waitKey(0)
                        gray6 = cv2.cvtColor(crop_image4, cv2.COLOR_BGR2GRAY)
                        if args["preprocess"] == "thresh":
                              gray6 = cv2.threshold(gray6, 0, 255,
                                    cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
                        elif args["preprocess"] == "blur":
                              gray6 = cv2.medianBlur(gray6, 3)
                        filename6 = "{}.png".format(os.getpid())
                        thresh4 = cv2.threshold(crop_image4, 150, 255, cv2.THRESH_BINARY_INV)[1]

                        result4 = cv2.GaussianBlur(thresh4, (5,5), 0)
                        result4 = 255 - result4

                        operator = pytesseract.image_to_string(result4, lang='eng',config='--psm 6')
                        opclose = difflib.get_close_matches(operator, operator_names, 1, 0.4)
                        print(opclose)
                        sheet1.write (x, 0, opclose)
                        x=x+1
                        #Operator 4-Blue
                        crop_image5 = image[478:45+478, 515+350+350:515+340+270+340]
                  
                        cv2.waitKey(0)
                        gray7 = cv2.cvtColor(crop_image5, cv2.COLOR_BGR2GRAY)
                        if args["preprocess"] == "thresh":
                              gray7 = cv2.threshold(gray7, 0, 255,
                                    cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
                        elif args["preprocess"] == "blur":
                              gray7 = cv2.medianBlur(gray7, 3)
                        filename7 = "{}.png".format(os.getpid())
                        thresh5 = cv2.threshold(crop_image5, 150, 255, cv2.THRESH_BINARY_INV)[1]

                        result5 = cv2.GaussianBlur(thresh5, (5,5), 0)
                        result5 = 255 - result5

                        operator = pytesseract.image_to_string(result5, lang='eng',config='--psm 6')
                        opclose = difflib.get_close_matches(operator, operator_names, 1, 0.4)
                        print(opclose)
                        sheet1.write (x, 0, opclose)
                        x=x+1
                        #Operator 5-Blue
                        crop_image6 = image[478:45+478, 515+350+350+350:515+340+270+340+340]
                  
                        cv2.waitKey(0)
                        gray8 = cv2.cvtColor(crop_image6, cv2.COLOR_BGR2GRAY)
                        if args["preprocess"] == "thresh":
                              gray8 = cv2.threshold(gray8, 0, 255,
                                    cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
                        elif args["preprocess"] == "blur":
                              gray8 = cv2.medianBlur(gray8, 3)
                        filename8 = "{}.png".format(os.getpid())
                        thresh6 = cv2.threshold(crop_image6, 150, 255, cv2.THRESH_BINARY_INV)[1]

                        result6 = cv2.GaussianBlur(thresh6, (5,5), 0)
                        result6 = 255 - result6

                        operator = pytesseract.image_to_string(result6, lang='eng',config='--psm 6')
                        opclose = difflib.get_close_matches(operator, operator_names, 1, 0.4)
                        print(opclose)
                        sheet1.write (x, 0, opclose)
                        x=x+1
                        #Operators Orange
                        #Operators 1-Orange
                        cropp_image2 = image[850:850+75, 175:175+270]
                  
                        cv2.waitKey(0)
                        grayy4 = cv2.cvtColor(cropp_image2, cv2.COLOR_BGR2GRAY)
                        if args["preprocess"] == "thresh":
                              grayy4 = cv2.threshold(grayy4, 0, 255,
                                    cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
                        elif args["preprocess"] == "blur":
                              grayy4 = cv2.medianBlur(grayy4, 3)
                        filenamee4 = "{}.png".format(os.getpid())
                        threshh2 = cv2.threshold(cropp_image2, 150, 255, cv2.THRESH_BINARY_INV)[1]

                        resultt2 = cv2.GaussianBlur(threshh2, (5,5), 0)
                        resultt2 = 255 - resultt2

                        operator = pytesseract.image_to_string(resultt2, lang='eng',config='--psm 6')
                        opclose = difflib.get_close_matches(operator, operator_names, 1, 0.4)
                        print(opclose)
                        sheet1.write (y, 2, opclose)
                        y = y +1

                        #Operators 2-Orange
                        cropp_image3 = image[850:850+75, 515:515+270]
                  
                        cv2.waitKey(0)
                        grayy5 = cv2.cvtColor(cropp_image3, cv2.COLOR_BGR2GRAY)
                        if args["preprocess"] == "thresh":
                              grayy5 = cv2.threshold(grayy5, 0, 255,
                                    cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
                        elif args["preprocess"] == "blur":
                              grayy5 = cv2.medianBlur(grayy5, 3)
                        filenamee5 = "{}.png".format(os.getpid())
                        threshh3 = cv2.threshold(cropp_image3, 150, 255, cv2.THRESH_BINARY_INV)[1]

                        resultt3 = cv2.GaussianBlur(threshh3, (5,5), 0)
                        resultt3 = 255 - resultt3

                        operator = pytesseract.image_to_string(resultt3, lang='eng',config='--psm 6')
                        opclose = difflib.get_close_matches(operator, operator_names, 1, 0.4)
                        print(opclose)
                        sheet1.write (y, 2, opclose)
                        y=y+1
                        #Operator 3-Orange
                        cropp_image4 = image[850:850+75, 515+350:515+340+270]

                        cv2.waitKey(0)
                        grayy6 = cv2.cvtColor(cropp_image4, cv2.COLOR_BGR2GRAY)
                        if args["preprocess"] == "thresh":
                              grayy6 = cv2.threshold(grayy6, 0, 255,
                                    cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
                        elif args["preprocess"] == "blur":
                              grayy6 = cv2.medianBlur(grayy6, 3)
                        filenamee6 = "{}.png".format(os.getpid())
                        threshh4 = cv2.threshold(cropp_image4, 150, 255, cv2.THRESH_BINARY_INV)[1]

                        resultt4 = cv2.GaussianBlur(threshh4, (5,5), 0)
                        resultt4 = 255 - resultt4

                        operator = pytesseract.image_to_string(resultt4, lang='eng',config='--psm 6')
                        opclose = difflib.get_close_matches(operator, operator_names, 1, 0.4)
                        print(opclose)
                        sheet1.write (y, 2, opclose)
                        y=y+1
                        #Operator 4-Orange
                        cropp_image5 = image[850:850+75, 515+350+350:515+340+270+340]

                        cv2.waitKey(0)
                        grayy7 = cv2.cvtColor(cropp_image5, cv2.COLOR_BGR2GRAY)
                        if args["preprocess"] == "thresh":
                              grayy7 = cv2.threshold(grayy7, 0, 255,
                                    cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
                        elif args["preprocess"] == "blur":
                              grayy7 = cv2.medianBlur(grayy7, 3)
                        filenamee7 = "{}.png".format(os.getpid())
                        threshh5 = cv2.threshold(cropp_image5, 150, 255, cv2.THRESH_BINARY_INV)[1]

                        resultt5 = cv2.GaussianBlur(threshh5, (5,5), 0)
                        resultt5 = 255 - resultt5

                        operator = pytesseract.image_to_string(resultt5, lang='eng',config='--psm 6')
                        opclose = difflib.get_close_matches(operator, operator_names, 1, 0.4)
                        print(opclose)
                        sheet1.write (y, 2, opclose)
                        y=y+1
                        #Operator 5-Orange
                        cropp_image6 = image[850:850+75, 515+350+350+350:515+340+270+340+340]
                  
                        cv2.waitKey(0)
                        grayy8 = cv2.cvtColor(cropp_image6, cv2.COLOR_BGR2GRAY)
                        if args["preprocess"] == "thresh":
                              grayy8 = cv2.threshold(grayy8, 0, 255,
                                    cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
                        elif args["preprocess"] == "blur":
                              grayy8 = cv2.medianBlur(grayy8, 3)
                        filenamee8 = "{}.png".format(os.getpid())
                        threshh6 = cv2.threshold(cropp_image6, 150, 255, cv2.THRESH_BINARY_INV)[1]

                        resultt6 = cv2.GaussianBlur(threshh6, (5,5), 0)
                        resultt6 = 255 - resultt6

                        operator = pytesseract.image_to_string(resultt6, lang='eng',config='--psm 6')
                        opclose = difflib.get_close_matches(operator, operator_names, 1, 0.4)
                        print(opclose)
                        sheet1.write (y, 2, opclose)
                        y=y+1
                        os.remove('./images/' + str(count) + ".jpg")
                        count = count + 90
                  
                  else:
                        os.remove('./images/' + str(count) + ".jpg")
                        count = count + 1
            else:
                        os.remove('./images/' + str(count) + ".jpg")
                        count = count + 1
      else:
            success == False

wb.save("./images/scanned/success.xls")