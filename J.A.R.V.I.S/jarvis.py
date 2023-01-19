import subprocess
import wolframalpha
import pyttsx3
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import pyjokes
import smtplib
import ctypes
import time
import requests
import shutil
import winshell

from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty("rate", 200)
engine.setProperty('voice', voices[1].id)
print(voices[1].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishme():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak('Good Morning sir!')
    elif hour >= 12 and hour < 17:
        speak('Good Afternoon sir!')
    elif hour >= 17 and hour < 20:
        speak('Good Evening sir!')
    else:
        speak('welcome sir')

    speak("I am JARVIS. How can i help you,Sir")

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('youremail@gmail.com', 'your-password')
    server.sendmail('youremail@gmail.com', to, content)
    server.close()


def takeCommand():
    #it takes microphonr input
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        r.energy_threshold = 500
        r.dynamic_energy_ratio = 2
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        print(e)
        print("Unable to Recognize your voice.")
        return "None"
    return query

    
if __name__ == "__main__":
    wishme()
    while True:
        query = takeCommand().lower()
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            speak("Opening Youtube")
            webbrowser.open("https://www.youtube.com/")

        elif 'open chrome' in query or 'open google' in query:
            speak("Here you go to Google")
            webbrowser.open("https://www.google.com/")

        elif 'open stack overflow' in query:
            speak("Opening Stackoverflow")
            webbrowser.open("https://www.stackoverflow.com/")   


        elif 'play music' in query or "play song" in query:
            music_dir = 'D:\\Music\\Kishore Kumar'
            songs = os.listdir(music_dir)
            num = random.randint(0, 27)    
            os.startfile(os.path.join(music_dir, songs[num]))

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")    
            speak(f"Sir, the time is {strTime}")

        elif 'open code' in query:
            codePath = "C:\\Users\\Anind\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)

        elif 'email to harry' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                to = "harryyourEmail@gmail.com"    
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("Sorry my friend. I am not able to send this email")

        elif 'go to sleep' in query:
            speak("Nice to meet you sir. Good Bye!")
            exit()
        
        elif 'who are you' in query:
            speak("Myself JARVIS, your personal voice assistant. I am created by Mr.Mitra. I can help you to simplify your task.")
        
        elif 'how are you' in query:
            speak("I am fine, Thank you")
            speak("How are you, Sir")
 
        elif 'fine' in query or "good" in query:
            speak("It's good to know that your fine")

        elif 'joke' in query:
            speak(pyjokes.get_joke())

        elif 'search' in query or 'play' in query:
             
            query = query.replace("search", "")
            query = query.replace("play", "")         
            webbrowser.open(f"https://www.google.com/search?q={query}")
 
        elif "who i am" in query:
            speak("If you talk then definitely your human.")

        elif 'lock window' in query:
                speak("locking the device")
                ctypes.windll.user32.LockWorkStation()
 
        elif 'shutdown system' in query:
                speak("Hold On a Sec ! Your system is on its way to shut down")
                subprocess.call('shutdown / p /f')
                 
        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
            speak("Recycle Bin cleared")

        elif "calculate" in query:
             
            app_id = "5VXHUY-4X86AV2AQG"
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('calculate')
            query = query.split()[indx + 1:]
            res = client.query(' '.join(query))
            answer = next(res.results).text
            print("The answer is " + answer)
            speak("The answer is " + answer)

        elif "what is" in query or "who is" in query:
        
            client = wolframalpha.Client("5VXHUY-4X86AV2AQG")
            res = client.query(query)
             
            try:
                print (next(res.results).text)
                speak (next(res.results).text)
            except StopIteration:
                print ("No results")

        elif "take a note" in query:
            speak("What should i write,sir")
            note = takeCommand()
            file = open('jarvis_notes.txt', 'w')
            speak("Sir, Should i include date and time")
            snfm = takeCommand()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("% H:% M:% S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)
            speak("Completed!")
         
        elif "the note" in query:
            speak("Showing Notes")
            file = open("jarvis_notes.txt", "r")
            print(file.read())
            speak(file.read())

        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("I am trying to Locate")
            speak(location)
            webbrowser.open(f"https://www.google.co.in/maps/place/{location}")
 
        elif "open camera" in query or "take a picture" in query:
            ec.capture(0, "Jarvis Camera ", "img.jpg")
        
        
