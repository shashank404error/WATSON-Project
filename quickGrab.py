from PIL import ImageGrab
import os
import time
import speech_recognition as sr   #use google speech recognition program 
from gtts import gTTS             #use google text-to-speech program 
from pygame import mixer          #use the mixer to play computer's speech
import win32api, win32con
from PIL import ImageOps
from PIL import ImageGrab
from numpy import *

x = 7
y = 85
def screenGrab():
   
    im = ImageGrab.grab((x,y,x+752,y+423))
    im.save(os.getcwd() + "\\full_snap_" + str(int(time.time())) + ".png", "PNG")
    
    print('screenShot captured')
#########################scanning image########################
def scan_img():
    box=(195,94,239,299)
    im=ImageOps.grayscale(ImageGrab.grab(box))
    a=array(im.getcolors())
    a=a.sum()
    
    print(a)
    
    return a

########################search start################

def startSearch():
    mousePos((534,34))
    leftClick()
    time.sleep(2)
    mousePos((225,66))
    leftClick()
    time.sleep(1)
    print('typing...')

#################type function######################


def startType(text):
    for i in range(0,len(text)):
        if text[i]=='a':
           mousePos((80,643))
           leftClick()
           time.sleep(0.3)
        elif text[i]=='b':
             mousePos((243,671))
             leftClick()
             time.sleep(0.3)
        elif text[i]=='c':
             mousePos((175,673))
             leftClick()
             time.sleep(0.3)
        elif text[i]=='d':
             mousePos((150,638))
             leftClick()
             time.sleep(0.3)
        elif text[i]=='e':
             mousePos((137,610))
             leftClick()
             time.sleep(0.3)
        elif text[i]=='f':
             mousePos((181,639))
             leftClick()
             time.sleep(0.3)
        elif text[i]=='g':
             mousePos((216,639))
             leftClick()
             time.sleep(0.3)
        elif text[i]=='h':
             mousePos((249,638))
             leftClick()
             time.sleep(0.3)
        elif text[i]=='i':
             mousePos((306,610))
             leftClick()
             time.sleep(0.3)
        elif text[i]=='j':
             mousePos((281,640))
             leftClick()
             time.sleep(0.3)
        elif text[i]=='k':
             mousePos((316,639))
             leftClick()
             time.sleep(0.3)
        elif text[i]=='l':
             mousePos((351,639))
             leftClick()
             time.sleep(0.3)
        elif text[i]=='m':
             mousePos((313,668))
             leftClick()
             time.sleep(0.3)
        elif text[i]=='n':
             mousePos((279,668))
             leftClick()
             time.sleep(0.3)
        elif text[i]=='o':
             mousePos((340,614))
             leftClick()
             time.sleep(0.3)
        elif text[i]=='p':
             mousePos((373,610))
             leftClick()
             time.sleep(0.3)
        elif text[i]=='q':
             mousePos((72,609))
             leftClick()
             time.sleep(0.3)
        elif text[i]=='r':
             mousePos((168,612))
             leftClick()
             time.sleep(0.3)
        elif text[i]=='s':
             mousePos((116,639))
             leftClick()
             time.sleep(0.3)
        elif text[i]=='t':
             mousePos((203,610))
             leftClick()
             time.sleep(0.3)
        elif text[i]=='u':
             mousePos((276,612))
             leftClick()
             time.sleep(0.3)
        elif text[i]=='v':
             mousePos((212,666))
             leftClick()
             time.sleep(0.3)
        elif text[i]=='w':
             mousePos((102,610))
             leftClick()
             time.sleep(0.3)
        elif text[i]=='x':
             mousePos((143,666))
             leftClick()
             time.sleep(0.3)
        elif text[i]=='y':
             mousePos((242,609))
             leftClick()
             time.sleep(0.3)
        elif text[i]=='z':
             mousePos((109,667))
             leftClick()
             time.sleep(0.3)
        elif text[i]=='.':
             mousePos((395,674))
             leftClick()
             time.sleep(0.3)
        elif text[i]=='/':
             mousePos((427,678))
             leftClick()
             time.sleep(0.3)
        else:
             mousePos((266,704))
             leftClick()
             time.sleep(0.3)

##########################enter button on keyboard#############
def enterbtn():
    time.sleep(0.2)
    mousePos((510,624))
    leftClick()
    

def enterbtn1():
    count=0
    while True:
        box1=(181,263,418,276)
        im=ImageOps.grayscale(ImageGrab.grab(box1))
        a=array(im.getcolors())
        a=a.sum()
        print(a)
        if(a==24322):
           mousePos((286,338))
           leftClick()
           break
        if(a==27647):
           mousePos((338,268))
           leftClick()
           break
        if(count==500):
           break
        count=count+1
           

################fb search functionalities######################
def fbsearchbox():
    count=0
    while True:
         box1=(15,86,462,121)
         im=ImageOps.grayscale(ImageGrab.grab(box1))
         a=array(im.getcolors())
         a=a.sum()
         print('fbsearchbox')
         print(a)
         if(a==43021):
            mousePos((187,108))
            leftClick()
            break
         if(count==250):
            break
         count=count+1

def openmsgoption():
    count=0
    while True:
         box1=(216,198,328,214)
         im=ImageOps.grayscale(ImageGrab.grab(box1))
         a=array(im.getcolors())
         a=a.sum()
         print('message option')
         print(a)
         if(a==6076):
            time.sleep(2)
            mousePos((668,330))
            leftClick()
            time.sleep(0.4)
            mousePos((639,366))
            leftClick()
            break
         else:
             box1=(571,203,692,225)
             im=ImageOps.grayscale(ImageGrab.grab(box1))
             b=array(im.getcolors())
             b=b.sum()
             print('message option for right name')
             print(b)
             if(b==6753):
                time.sleep(2)
                mousePos((671,255))
                leftClick()
                break
         if(count==100):
            break
         count=count+1

def sendmsg1():
    mousePos((624,290))
    leftClick()
    

def sendbtn():
    mousePos((564,403))
    leftClick()
    

def lookformsgbox():
    count=0
    while True:
          box1=(158,131,600,166)
          im=ImageOps.grayscale(ImageGrab.grab(box1))
          b=array(im.getcolors())
          b=b.sum()
          print('looking for message box')
          print(b)
          if(b==20674):
             print('opened message box')
             print('typing message')
             break
          if(count==500):
              break
          count=count+1


############################for image recognition#######################    
def openlink():
    box1=(379,579,522,685)
    im=ImageOps.grayscale(ImageGrab.grab(box1))
    b=array(im.getcolors())
    b=b.sum()
    print('from mobile cam')
    print(b)
    im.save(os.getcwd() + "\\full_snap_" + str(int(time.time())) + ".png", "PNG")
    return b

#################################fb search meta function#######################
def fb_msg(send_to,msg):
    startSearch()
    startType('facebook')
    enterbtn()
    enterbtn1()
    fbsearchbox()
    startType(send_to)
    enterbtn()
    openmsgoption()
    sendmsg1()
    lookformsgbox()
    startType(msg)
    sendbtn()
        
##########################wikipedia search #####################
def wiki_search(search_obj):
    startSearch()
    startType(search_obj)
    startType(' w')
    enterbtn()
    enterbtn1()
#######################avg processing image#########################
def avg_cam_val():
    cam1=[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,111,212,31,41,51,61,71,81,91,101,111,121,131,141,151,161,171,181,191,201,211,221,231,241,251]
    cam_sum=0
    super_sum=0
    cam_avg=[1,2]
    for j in range(0,2):
        for i in range(0,50):
             cam1[i]=scan_img()
             cam_sum=cam_sum+cam1[i]
        cam_avg[j]=(cam_sum)/50
        super_sum=super_sum+cam_avg[j]
    super_avg=(super_sum)/2 
    print(super_avg)

######################trigger fb msg#################
def trig_fb_msg():
    count=0
    while True:
        box1=(233,391,260,403)
        im=ImageOps.grayscale(ImageGrab.grab(box1))
        a=array(im.getcolors())
        a=a.sum()
        print('comnmand sent from mob')
        print(a)
        if(a==2590):
             print('fb_msg_initialised')
             time.sleep(0.7)
             mousePos((213,441))
             leftClick()
             startType('function initiated')
             enterbtn()
             fb_msg('ashish kumar','yo bro')
             time.sleep(4)
             mousePos((213,441))
             leftClick()
             startType('message sent')
             enterbtn()
             break
        if(count==700):
              break
        count=count+1
             
###################start seeing##################
def watch_it():
    cam1=scan_img()
    if (cam1>39000 and cam1<42100):
        obj_found='charan ka bottle'
    elif (cam1>20000 and cam1<24000):
        obj_found='blue bottle'
    elif (cam1>35000 and cam1<37000):
        obj_found='purple_bottle'
    elif (cam1>38000 and cam1<40000):
        obj_found='coffee_mug'
    else:
        obj_found='nothing'

    return obj_found
    
#######################say it##########################

def play_music(file_name):
    mixer.init()   # start mixer to play mp3
    mixer.music.set_volume(1)
    mixer.music.load(file_name) # load file
    mixer.music.play() # play file
    
    
    
#################mouse controls###################
    
def leftClick():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0)
    time.sleep(1)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0)
    

def leftDown():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0)
    time.sleep(.1)
    print('left down')

def leftUp():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0)
    time.sleep(.1)
    print('left release')

def mousePos(cord):
    win32api.SetCursorPos((cord[0],cord[1]))

def get_cords():
    x,y=win32api.GetCursorPos()
    print(x,y)

###################start telling########################
count=0
while False:
    if (count==5):
        break      
    else:
        obj_in=watch_it()
        print(obj_in)
        
        time.sleep(5)
    count=count+1
    



if __name__ == '__main__':
   trig_fb_msg()
       
