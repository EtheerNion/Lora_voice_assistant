from datetime import datetime
import datetime
import smtplib
from neuralintents import GenericAssistant
import speech_recognition
import pyttsx3 as tts
import win32com.client
import sys
import wikipedia
import webbrowser
import pyautogui
import pyowm
import os
import ctypes
import subprocess


recognizer = speech_recognition.Recognizer()


speaker = tts.init()
speaker.setProperty('rate', 150)
speaker.setProperty('voice', '0x041f')


todo_list = []


def create_note():
    global recognizer

    speaker.say("What do you want to note?")
    speaker.runAndWait()

    done = False

    while not done:
        try:

            with speech_recognition.Microphone() as mic:

                recognizer.adjust_for_ambient_noise(mic, duration=0.2)
                audio = recognizer.listen(mic)

                note = recognizer.recognize_google(audio)
                note = note.lower()

                speaker.say("Choose a file name")
                speaker.runAndWait()
                
                recognizer.adjust_for_ambient_noise(mic, duration=0.2)
                audio = recognizer.listen(mic)

                filename = recognizer.recognize_google(audio)
                filename = filename.lower()

            with open(filename, 'w') as f:
                f.write(note)
                done = True
                speaker.say("note successfully created {}".format(filename)) #filename used with .format()
                speaker.runAndWait()


        except speech_recognition.UnknownValueError:
            recognizer = speech_recognition.Recognizer()
            speaker.say("Can not understand")
            speaker.runAndWait()


def add_todo():
    global recognizer

    speaker.say("What do you want to add in to-do?")
    speaker.runAndWait()

    done = False 

    while not done:
        try:

            with speech_recognition.Microphone() as mic:

                recognizer.adjust_for_ambient_noise(mic, duration=0.2)
                audio = recognizer.listen(mic)

                item = recognizer.recognize_google(audio)
                item = item.lower()

                todo_list.append(item)
                done = True

                speaker.say("{item} succesfuly added")
                speaker.runAndWait()

        except speech_recognition.UnknownValueError:
            recognizer = speech_recognition.Recognizer()
            speaker.say("Can not understand")
            speaker.runAndWait()


def list_todo():
    speaker.say("These are in to-do")

    for item in todo_list:
        speaker.say(item)
    speaker.runAndWait()


def greeting():
    speaker.say("Hello, how can I help you?")
    speaker.runAndWait()


def goodbye():
    speaker.say("Bye")
    speaker.runAndWait()
    sys.exit(0)


def what_time():
    now = datetime.now()
    s = now.strftime("%H %M %S")
    speaker.say("Time is", s)
    speaker.runAndWait()


def search_in_wiki():

    global recognizer

    speaker.say("What do you want to search in Wiki?")
    speaker.runAndWait()

    done = False 

    while not done:
        try:

            with speech_recognition.Microphone() as mic:

                recognizer.adjust_for_ambient_noise(mic, duration=0.2)
                audio = recognizer.listen(mic)

                wiki = recognizer.recognize_google(audio)
                wiki = wiki.lower()
                results = wikipedia.summary(wiki, sentences = 3) # wiki used as title
                done = True

                speaker.say("{wiki} I found in wiki ")
                speaker.say(results)
                speaker.runAndWait()

        except speech_recognition.UnknownValueError:
            recognizer = speech_recognition.Recognizer()
            speaker.say("Can not understand")
            speaker.runAndWait()


def open_video_youtube():

    global recognizer

    speaker.say("I am looking for the video")
    speaker.runAndWait()

    done = False 

    while not done:
        try:

            with speech_recognition.Microphone() as mic:

                recognizer.adjust_for_ambient_noise(mic, duration=0.2)
                audio = recognizer.listen(mic)

                VideoName = recognizer.recognize_google(audio)
                VideoName = VideoName.lower()

                adres = ('https://www.youtube.com/results?search_query=')+VideoName
                webbrowser.get().open(adres)
                pyautogui.moveTo(x=702, y=252)
                pyautogui.click(button='left', clicks=2, interval=7)                      
             
        except speech_recognition.UnknownValueError:
            recognizer = speech_recognition.Recognizer()
            speaker.say("Can not understand")
            speaker.runAndWait()

        #exit page with pyautogui
        pyautogui.hotkey('ctrl', 'w') 
      

def search_google():

    global recognizer

    speaker.say("I am looking for it")
    speaker.runAndWait()

    done = False 

    while not done:
        try:

            with speech_recognition.Microphone() as mic:

                recognizer.adjust_for_ambient_noise(mic, duration=0.2)
                audio = recognizer.listen(mic)

                search = recognizer.recognize_google(audio)
                search = search.lower()

                adres = ('https://www.google.com/search?q=')+search
                webbrowser.get().open(adres)
                pyautogui.moveTo(x=384, y=368)
                pyautogui.click(button='left',clicks=2,interval=7)  

                speaker.say("I found in google for {arama}")
                speaker.runAndWait()                   
             
        except speech_recognition.UnknownValueError:
            recognizer = speech_recognition.Recognizer()
            speaker.say("Can not understand")
            speaker.runAndWait()


def open_twitter():
    
    adres = ('https://www.twitter.com/')
    webbrowser.open_new_tab(adres)


def open_github():

    adres = ('https://github.com/')
    webbrowser.open_new_tab(adres)


def day_week():
    
    day = datetime.datetime.now()
    speaker.say(datetime.datetime.strftime(day, '%A'))
    speaker.runAndWait()


def current_date():

    current_date = datetime.datetime.now()
    speaker.say(datetime.datetime.strftime(current_date, '%d %B %Y'))
    speaker.runAndWait()


def weather_forc():
    global recognizer

    owm = pyowm.OWM("e402c33c5ae91a07f0adc05ef341c1c5")
    mgr = owm.weather_manager()

    speaker.say("For which city you would like to see weather forecast?")
    speaker.runAndWait()
#city issue
    done = False 

    while not done:
        try:

            with speech_recognition.Microphone() as mic:

                recognizer.adjust_for_ambient_noise(mic, duration=0.2)
                audio = recognizer.listen(mic)

                city = recognizer.recognize_google(audio)
                city = city.lower()

                #first letter of the city in capital
                city = city.capitalize() 

                forcast = mgr.weather_at_place("{city},EU")
                hava = forcast.weather

                details = hava.get_detailed_status()
                degree = hava.get_temperature('celcius')["temp"]

                speaker.say(" for {city} weather forecast {details} and {degree} degree.")
             
        except speech_recognition.UnknownValueError:
            recognizer = speech_recognition.Recognizer()
            speaker.say("Can not understand")
            speaker.runAndWait()


def create_email():
    global recognizer

    speaker.say("Who do you want to send an email?")
    speaker.runAndWait()

    done = False 

    while not done:
        try:

            with speech_recognition.Microphone() as mic:

                recognizer.adjust_for_ambient_noise(mic, duration=0.2)
                audio = recognizer.listen(mic)

                username = input ("Email Address")
                password = input ("Password")
#cikis yapti
                server = smtplib.SMTP('smtp.gmail.com', 587)
                server.ehlo()
                server.starttls()
                server.login("{username}", "{password}")
                to = ["Email text"]
                email_body ="test"
                try:
                    server.sendmail(username, to, email_body)
                    print("Email successfully sent")
                except:
                    print("Email unsuccessful")

                server.quit()

        except speech_recognition.UnknownValueError:
            recognizer = speech_recognition.Recognizer()
            speaker.say("Can not understand")
            speaker.runAndWait()


def read_email():

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items
    message = messages.GetLast()

    speaker.say("there is an email from {message.SenderName} , email title {message.subject}, index {message.body}")
    speaker.runAndWait()


def open_folder():
    
    speaker.say("What do you want me to open?")
    speaker.runAndWait()

    done = False 

    while not done:
        try:

            with speech_recognition.Microphone() as mic:

                recognizer.adjust_for_ambient_noise(mic, duration=0.2)
                audio = recognizer.listen(mic)

                folder = recognizer.recognize_google(audio)
                folder = folder.lower()

                os.system('cd ${folder}')
             
        except speech_recognition.UnknownValueError:
            recognizer = speech_recognition.Recognizer()
            speaker.say("can not understand")
            speaker.runAndWait()


def create_folder():

    global recognizer

    speaker.say("Where the folder be created?")
    speaker.runAndWait()

    done = False 

    while not done:
        try:

            with speech_recognition.Microphone() as mic:

                recognizer.adjust_for_ambient_noise(mic, duration=0.2)
                audio = recognizer.listen(mic)

                folder_place = recognizer.recognize_google(audio)
                folder_place = folder_place.lower()

                if folder_place == "here" or "there" or "bu dizin":
                    os.mkdir()

                else:
                    os.makedirs("{folder_place}", mode = 0o755, exist_ok=True)

        except speech_recognition.UnknownValueError:
            recognizer = speech_recognition.Recognizer()
            speaker.say("Can not understand")
            speaker.runAndWait()


def system_shut_down():

    speaker.say("Do you want to shot down your PC?")
    speaker.runAndWait()

    done = False 

    while not done:
        try:

            with speech_recognition.Microphone() as mic:

                recognizer.adjust_for_ambient_noise(mic, duration=0.2)
                audio = recognizer.listen(mic)

                shutdown = recognizer.recognize_google(audio)
                shutdown = shutdown.lower()

                if shutdown == 'no':
                   exit()
                else:
                    os.system("shutdown /s /t 1")           
             
        except speech_recognition.UnknownValueError:
            recognizer = speech_recognition.Recognizer()
            speaker.say("Can not understand")
            speaker.runAndWait()
            

def create_alarm():

    print("Alarm is being created")
  
    
def list_alarms():
    
    print("Alarm list")


def antivirus():
    speaker.say("Starting virus scanning")
    subprocess.Popen('powershell.exe [Start-MpScan')


def lock_screen():

    speaker.say("locking the device")
    ctypes.windll.user32.LockWorkStation()
    


mappings = {

    'greeting': greeting,
    'create_note': create_note,
    'add_todo': add_todo,
    'list_todo': list_todo,
    'goodbye': goodbye,
    'what_time': what_time,
    'search_in_wiki': search_in_wiki,
    'open_video_youtube': open_video_youtube,
    'search_google': search_google,
    'open_twitter': open_twitter,
    'open_github': open_github,
    'day_week': day_week,
    'current_date': current_date,
    'weather_forc': weather_forc,
    'create_email': create_email,
    'read_email': read_email,
    'open_folder': open_folder,
    'create_folder': create_folder,
    'system_shut_down': system_shut_down,
    'create_alarm': create_alarm,
    'list_alarms': list_alarms,
    'antivirus': antivirus,
    'lock_screen': lock_screen
}  


assistant = GenericAssistant('intents.json', intent_methods=mappings)
assistant.train_model()

while True:
    try:
        with speech_recognition.Microphone() as mic:

            recognizer.adjust_for_ambient_noise(mic, duration=0.2)
            audio = recognizer.listen(mic)

            message = recognizer.recognize_google(audio)
            message = message.lower()

        assistant.request(message)

    except speech_recognition.UnknownValueError:
        recognizer = speech_recognition.Recognizer()
