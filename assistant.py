import subprocess
import wolframalpha
import pyttsx3
import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import time
import requests
import shutil
from twilio.rest import Client
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
import webbrowser
import os

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>= 0 and hour<12:
        speak("Good Morning Sir !")

    elif hour>= 12 and hour<18:
        speak("Good Afternoon Sir !") 

    else:
        speak("Good Evening Sir !") 

    assname =("Bot")
    speak("I am your Assistant")
    speak(assname)
    

def username():
    speak("What should i call you sir")
    uname = takeCommand()
    speak("Welcome Mister")
    speak(uname)
    columns = shutil.get_terminal_size().columns
    
    print("#####################".center(columns))
    print("Welcome Mr.", uname)
    print("#####################".center(columns))
    
    speak("How can i Help you, Sir")

def takeCommand():
    
    r = sr.Recognizer()
    
    with sr.Microphone() as source:
        
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...") 
        query = r.recognize_google(audio, language ='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        print(e) 
        print("Unable to Recognize your voice.") 
        return "None"
    
    return query

if __name__ == '__main__':
    clear = lambda: os.system('cls')
    
    # This Function will clean any
    # command before execution of this python file
    clear()
    wishMe()
    username()
    
    while True:
        
        query = takeCommand().lower()
        
        # All the commands said by user will be 
        # stored here in 'query' and will be
        # converted to lower case for easily 
        # recognition of command
        if 'search' in query:
            if 'google' in query:
                speak('Searching google...')
                a=query.find('in google')
                if a !=-1:
                    query=query[:a].strip()
                search_url=f"https://www.google.com/search?q={query}"
                webbrowser.open(search_url)
            elif 'wikipedia' in query:
                speak('Searching Wikipedia...')
                query = query.replace("wikipedia", "")
                results = wikipedia.summary(query, sentences = 3)
                speak("According to Wikipedia")
                print(results)
                speak(results)
            elif 'news' in query:
                try: 
                    jsonObj = urlopen('''https://newsapi.org / v1 / articles?source = the-times-of-india&sortBy = top&apiKey =apikey''')
                    data = json.load(jsonObj)
                    i = 1
                
                    speak('here are some top news from the times of india')
                    print('''=============== TIMES OF INDIA ============'''+ '\n')
                
                    for item in data['articles']:
                    
                        print(str(i) + '. ' + item['title'] + '\n')
                        print(item['description'] + '\n')
                        speak(str(i) + '. ' + item['title'] + '\n')
                        i += 1
                except Exception as e:
                
                    print(str(e))
        elif 'calculate' in query:
            app_id = "api key"
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('calculate') 
            query = query.split()[indx + 1:] 
            res = client.query(' '.join(query)) 
            answer = next(res.results).text
            print("The answer is " + answer) 
            speak("The answer is " + answer)
        elif 'open' in query:
            if 'browser' in query or 'google' in query:
                speak("Here you go to Google\n")
                webbrowser.open("google.com")
            elif 'youtube' in query :
                speak("Here you go to Youtube\n")
                webbrowser.open("youtube.com")
            elif 'gmail' in query:
                speak("Here you go to gmail\n")
                webbrowser.open("gmail.com")
        elif 'music' in query or 'play song' in query:
            speak("Here you go with music")
            # music_dir = "G:\\Song"
            music_dir = "C:\\Users\\Admin\\Music"         # Change this directory path according to your system
            songs = os.listdir(music_dir)
            print(songs) 
            random = os.startfile(os.path.join(music_dir, songs[0]))
        elif 'system' in query:
            if 'lock' in query:
                speak("locking the device")
                ctypes.windll.user32.LockWorkStation()
            elif 'shutdown' in query:
                speak("Hold On a Sec ! Your system is on its way to shut down")
                subprocess.call('shutdown / p /f')
            elif 'empty recycle bin' in query:
                winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
                speak("Recycle Bin Recycled")
            elif 'restart' in query:
                subprocess.call(["shutdown", "/r"])
            elif 'log off' in query or 'sign out' in query:
                speak("Make sure all the application are closed before sign-out")
                time.sleep(5)
                subprocess.call(["shutdown", "/l"])
        elif 'change' in query:
            if 'my name' in query:
                query = query.replace("change my name to", "")
                assname = query
            elif 'your name' in query:
                speak("What would you like to call me, Sir ")
                assname = takeCommand()
                speak("Thanks for naming me")
        elif 'time'in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S") 
            speak(f"Sir, the time is {strTime}")
        elif 'about you' in query or 'about yourself' in query:
            speak("I am a virtual assitant, created to help you")
        elif 'exit' in query:
            speak("Thanks for giving me your time")
            exit()
        elif 'joke' in query:
            speak(pyjokes.get_joke())
        elif "don't listen" in query or "stop listening" in query:
            speak("for how much time you want to stop jarvis from listening commands")
            a = int(takeCommand())
            time.sleep(a)
            print(a)
        elif "weather" in query:
            
            # Google Open weather website
            # to get API of Open weather 
            api_key = "api key"
            base_url = "http://api.openweathermap.org / data / 2.5 / weather?"
            speak(" City name ")
            print("City name : ")
            city_name = takeCommand()
            complete_url = base_url + "appid =" + api_key + "&q =" + city_name
            response = requests.get(complete_url) 
            x = response.json() 
            
            if x["code"] != "404": 
                y = x["main"] 
                current_temperature = y["temp"] 
                current_pressure = y["pressure"] 
                current_humidiy = y["humidity"] 
                z = x["weather"] 
                weather_description = z[0]["description"] 
                print(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidiy) +"\n description = " +str(weather_description)) 
            
            else: 
                speak(" City Not Found ")
        elif "Good Morning" in query:
            speak("A warm" +query)
            speak("How are you Mister")
            speak(assname)
        else:
            speak("Query was not understood")
