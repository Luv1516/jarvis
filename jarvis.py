import speech_recognition as sr
import pyttsx3
import os
import webbrowser
import datetime
import random
import win32com.client

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voices', voices[0].id)

def say(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)

    if hour>0 and hour<12:
        say("Good Morning!")

    elif hour>=12 and hour<16:
        say("Good Afternoon!")
         
    else:
        say("Good Evening!")   

    say("I am Jarvis A.I., what can I help you with,sir") 

def takeCommand():
    r = sr.Recognizer()

    with sr.Microphone() as source:
        print("listening...")
        audio = r.listen(source)
        try:
            print("recognizing...")
            query = r.recognize_google(audio,language="en-in")   
            print(f"user said: {query}")     
            return query
        except Exception as e:
            return "sorry , some error has occurred."

if __name__ == '__main__':
    wishMe()    

    while True:
        query = takeCommand()
        if "open youtube".lower() in query.lower():
            say("opening youtube")
            webbrowser.open("https://www.youtube.com")

        if "open google".lower() in query.lower():
            say("opening google")
            webbrowser.open("https://www.google.com") 

        if "open Chess.com".lower() in query.lower():
            say("opening chess.com")
            webbrowser.open("https://www.chess.com/home",)  

        if 'play music'.lower() in query.lower():
            music_dir=  'C:\\Users\\LUV\\Music'
            songs = os.listdir(music_dir)
            os.startfile(os.path.join(music_dir,songs[random.randrange(152)]))  

        if 'jarvis do you know me'.lower() in query.lower():
            say("yes Sir, you are my creator , mister Luv ")

        if "rick roll me".lower() in query.lower():   
            say("there you go") 
            webbrowser.open("https://www.youtube.com/watch?v=xvFZjo5PgG0")

        if "time".lower() in query.lower(): 
            strfTime = datetime.datetime.now().strftime("%H hours and %M minutes") 
            say(f"the time is{strfTime}")

        if "good".lower() in query.lower():
            say("thanks! I appreciate that")

        if "visual studio".lower() in query.lower():
            code = 'C:\\Users\\LUV\\Desktop\\VSCode-win32-x64-1.84.2\\Code'
            say("opening visual studio")
            os.startfile(code)    

        if "open notepad".lower() in query.lower():
            say("opening notepad")
            os.startfile("C:\\Windows\\System32\\notepad.exe")
            
        if "open word".lower() in query.lower():
            say("opening word")
            os.startfile("E:\\Office 2013\\Software Files\\Office15\\WINWORD.EXE")

        if "open excel".lower() in query.lower():
            say("opening excel")
            os.startfile("E:\\Office 2013\\Software Files\\Office15\\EXCEL.EXE")

        if "open powerpoint".lower() in query.lower():
            say("opening powerpoint")
            os.startfile("E:\\Office 2013\\Software Files\\Office15\\POWERPNT.EXE")

        if "thanks".lower() in query.lower():
            say("you are welcome")
            
        if "hello jarvis".lower() in query.lower():
            say("hello sir, what can I help you with?")

        if "bye".lower() in query.lower():
            say("bye sir")
            break    
