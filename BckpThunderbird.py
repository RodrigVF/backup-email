#----------------------------IMPORTS---------------------------
import os
import pyautogui
pyautogui.FAILSAFE = False
import time 
import subprocess
from screeninfo import get_monitors
import ctypes
from dotenv import load_dotenv

load_dotenv()
THUNDERBIRD_PASSWORD = os.getenv('THUNDERBIRD_PASSWORD')
THUNDERBIRD_IMAP = os.getenv('THUNDERBIRD_IMAP')
THUNDERBIRD_IMAP_PORT = os.getenv('THUNDERBIRD_IMAP_PORT')
THUNDERBIRD_SMTP = os.getenv('THUNDERBIRD_SMTP')
THUNDERBIRD_SMTP_PORT = os.getenv('THUNDERBIRD_SMTP_PORT')
THUNDEBIRD_APP = os.getenv('THUNDEBIRD_APP')
THUNDEBIRD_CACHE = os.getenv('THUNDEBIRD_CACHE')
BACKUP_FOLDER = os.getenv('BACKUP_FOLDER')
EMAIL_DOMAIN = os.getenv('EMAIL_DOMAIN')


#-------------------------GET SCREEN INFO------------------------
#Get Scale
scaleFactor = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100
print(scaleFactor)
#Get Resolution
for monitor in get_monitors():
  
  if (monitor.is_primary == True):
    screen_width = monitor.width
    screen_height = monitor.height


resolutionXvariable = (screen_width / 1920) 
resolutionYvariable = (screen_height / 1080)

#----------------------------FUNCTIONS-------------------------
def tab_sequence(n):
  counter = 1
  while (counter <= n):
    pyautogui.press('tab')
    counter += 1
  time.sleep(1) 


def wait_load(x,y,rgb):
    newX = int(x*scaleFactor)
    newY = int(y*scaleFactor)
    checkImage = False
    while (checkImage == False):
      checkImage = pyautogui.pixelMatchesColor(newX,newY,rgb,tolerance=10)
      if(rgb == (18,188,0)):
        pyautogui.click(1,65*scaleFactor)
        pyautogui.hotkey('ctrl', 'home')
      time.sleep(2) 
      

def change_auth_method():
    pyautogui.press('space')
    time.sleep(1) 
    pyautogui.press('up')
    pyautogui.press('up')
    time.sleep(1) 
    pyautogui.press('enter')
    time.sleep(1) 
    
#--------------------------CREATE LIST------------------------
userArray = []
username = ''
print('After send all users you want to make backup, type "done" and press enter')
while (username != 'done'):
  username = input('User (ex: user-email):\n')
  if (username != 'done'):
    userArray.append(username)
#----------------------------START----------------------------
while True:
    if (len(userArray) == 0):
      subprocess.call(["cmd", "/c", "start", "/max", BACKUP_FOLDER]) 
      time.sleep(2)
      break
#----------------------------VARIABLES------------------------
    userEmail = userArray[0]+EMAIL_DOMAIN
    splitUserFullName = userArray[0].split('-')
    userFullName = ' '.join(splitUserFullName).title()
    checkImage = False
    counter = 1
  #--------------------------LOGIN PAGE-----------------------
    print('Starting backup of: '+userFullName)
  #Open/Close ThunderBird
    subprocess.call(["cmd", "/c", "start", "/max", THUNDEBIRD_APP])
    wait_load(727,331, (254,254,254))
    time.sleep(.5)
    pyautogui.hotkey('alt', 'f4')
  
  #Open ThunderBird
    time.sleep(1)
    subprocess.call(["cmd", "/c", "start", "/max", THUNDEBIRD_APP])
    time.sleep(3)
    wait_load(727,331, (254,254,254))
    time.sleep(1)
  #Erase default username
    pyautogui.keyDown('shiftleft')
    pyautogui.keyDown('shiftright')
    time.sleep(1) 
    pyautogui.press('home')
    pyautogui.keyUp('shiftleft')
    pyautogui.keyUp('shiftright')
    pyautogui.press('backspace')
    time.sleep(1) 
    
  #Write user fullname
    pyautogui.write(userFullName)
    
  #Write email
    pyautogui.press('tab')
    time.sleep(1) 
    pyautogui.write(userEmail)
    
  #Write default password
    pyautogui.press('tab')
    time.sleep(1) 
    pyautogui.write(THUNDERBIRD_PASSWORD)
    
  #Press Login
    tab_sequence(5)
    pyautogui.press('enter')
  
  #--------------------SERVER CONNECT PAGE--------------------
  #Wait Load
    time.sleep(4) 
    wait_load(350,470, (18,188,0))
      
  #Go to "configurar manualmente"
    tab_sequence(7)
    pyautogui.press('enter')
    time.sleep(1) 
    
  #Configure IMAP/SMTP
    pyautogui.press('tab')
    
  #Type IMAP
    time.sleep(1) 
    pyautogui.write(THUNDERBIRD_IMAP)
    pyautogui.press('tab') 
    
  #Type IMAP Port
    time.sleep(1) 
    pyautogui.write(THUNDERBIRD_IMAP_PORT)
    tab_sequence(3)
    
  #Type e-mail
    time.sleep(1) 
    pyautogui.write(userEmail)
    pyautogui.press('tab')
    
  #Type SMTP
    time.sleep(1) 
    pyautogui.write(THUNDERBIRD_SMTP)
    time.sleep(1) 
    pyautogui.press('tab') 
  #Type Port
    time.sleep(1) 
    pyautogui.write(THUNDERBIRD_SMTP_PORT)
    tab_sequence(5)
  #Test e-mail
    pyautogui.press('enter')
    time.sleep(1) 
    
  #Wait test
    wait_load(361,468, (18,188,0))
  #Go to auth method
    tab_sequence(10)
    
  #Change auth method
    change_auth_method()
    
    pyautogui.press('tab')
  
  #Go to auth method
    tab_sequence(4)
    
  #Change auth method
    change_auth_method()
    
  #Press ok
    tab_sequence(5)
    pyautogui.press('enter')
    
  #------------------USER PREFERENCES PAGE-------------------
  #Wait Load page
    time.sleep(2) 
    
    wait_load(825,350, (0,179,244))
    time.sleep(1) 
  
  #Press ok
    tab_sequence(8)
    pyautogui.press('enter')
  
  #Wait last load
    wait_load(147,286, (56,56,61))
    time.sleep(1) 
  
  #Press ok
    tab_sequence(3)
    pyautogui.press('enter')
    
  #---------------------WAIT BACKUP--------------------------
  #blue bar  = 1525,1003 69,177,255 #45B1FF 
  # 2294,1017 69,177,255 #45B1FF
  # 1640 || 2290
  # 2208,1017 69,177,255 #45B1FF ---- ADD LATER
  #Bug bar = 2240,1017 84,183,255 #54B7FF
    print('Backup Started!')
    time.sleep(15) 

    X1 = int(int(screen_width)-(302*scaleFactor))
    X2 = int(int(screen_width)-(294*scaleFactor))
    X3 = int(int(screen_width)-(289*scaleFactor))
    x4 = int(int(screen_width)-(282*scaleFactor))
    bugX = int(int(screen_width)-(256*scaleFactor))
    newY = int(int(screen_height)-(50*scaleFactor))
    
    print(X1)
    print(X2)
    print(X3)
    print(bugX)
    print(newY)
    
    #Check backup bar
    backupCounter = 7
    while (backupCounter != 0):
      time.sleep(5)
      pixel1 = pyautogui.pixelMatchesColor(X1,newY, (69,177,255), tolerance=10)
      pixel2 = pyautogui.pixelMatchesColor(X2,newY, (69,177,255), tolerance=10)
      pixel3 = pyautogui.pixelMatchesColor(X3,newY, (69,177,255), tolerance=10)
      pixelBug = pyautogui.pixelMatchesColor(bugX,newY, (84,183,255), tolerance=30)

      if ( pixel1 == False and pixel2 == False and pixel3 == False and pixelBug == False): 
        messageBackupCheck = 'Backup Stoped?'
        print(f'{messageBackupCheck}({backupCounter})')
        backupCounter -= 1
        time.sleep(15) 

    pyautogui.hotkey('alt', 'f4')
    time.sleep(2)
    
        
  #---------------------SAVE BACKUP--------------------------
    #Cut from app  
    subprocess.call(["cmd", "/c", "start", "/max", THUNDEBIRD_CACHE])  
    time.sleep(2)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'x')
    time.sleep(1)
    #Open Backup Folder
    subprocess.call(["cmd", "/c", "start", "/max", BACKUP_FOLDER]) 
    time.sleep(2)
    pyautogui.hotkey('ctrl','shift', 'n')
    time.sleep(1)
    pyautogui.write(userArray[0])
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    #Paste to User Backup Folder
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'v')
    print('Finishing backup of: '+userArray[0])
    
    #------------------REMOVE USER FROM LIST-----------------
    time.sleep(2) 
    pyautogui.hotkey('alt', 'f4')
    time.sleep(1) 
    print('wait...')
    userArray.pop(0)
    time.sleep(3) 
    
    #break
    
    #Check if thunderbird glitched
    if (pyautogui.pixelMatchesColor(bugX,newY, (84,183,255), tolerance=30) == True):
      messageBackupCheck = 'Loading... '
      time.sleep(20)
      if (pyautogui.pixelMatchesColor(bugX,newY, (84,183,255), tolerance=30) == False): 
        pyautogui.hotkey('alt', 'f4')
        time.sleep(2)
        subprocess.call(["cmd", "/c", "start", "/max", BACKUP_FOLDER]) 
        time.sleep(15) 