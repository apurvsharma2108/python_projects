import pyttsx3
import datetime
import random
import speech_recognition as sr 
import wikipedia
import psutil
import webbrowser
import os
import time
from win32com.client import Dispatch
import requests
import json
import bs4

engine=pyttsx3.init('sapi5')
engine.setProperty('rate',175)
engine.setProperty('volume',0.9)
voices=engine.getProperty('voices')   #collects voices
engine.setProperty('voice',voices[1].id)
def speak(audio):
    engine.say(audio)
    engine.runAndWait()
def takeCommand():
    #It takes microphone input from the user and returns string output
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:   
        print("recognizing")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        print(e)    
        print("Say that again please...")  
        return "None"
    return query

def wish_me():
    hour=int(datetime.datetime.now().hour)
    if(hour>6 and hour<=12):
        speak("good morning sir i am friday please tell me how may i help you")
    elif(hour>12 and hour<18):
        speak("good afternoon sir i am friday please tell me  how may i help you")
    else:
        speak("good evening sir i am friday please tell me  how may i help you ")

if __name__ == "__main__":
    wish_me()
    while True:
        query=takeCommand().lower()
        if 'wikipedia' in query:
            speak("searching wikipedia")
            query=query.replace("wikipedia","")
            results=wikipedia.summary(query,sentences=2)
            speak("according to wikipedia")
            print(results)
            speak(results)
        elif 'youtube' in query:
            speak("opening youtube")
            webbrowser.open("youtube.com")
        elif 'play music' in query:
            speak("playing music")
            music_dir='D:\\music'
            songs=os.listdir(music_dir)
            s=random.randint(0,2)
            os.startfile(os.path.join(music_dir,songs[s]))

        elif 'time' in query:
             strTime = datetime.datetime.now().strftime("%H:%M:%S")    
             speak(f"the time is {strTime}")
        
        elif 'latest news' in query:
            speak("telling the latest news")
            url = ('http://newsapi.org/v2/top-headlines?country=in&apiKey=4f1a04ef66b642fe8791dda97b749170')
            response = requests.get(url).text
            news_dict = json.loads(response)
            articles = news_dict['articles']
            for arts in articles:
                speak(arts['title'])
                speak("Moving on to next news....")
            speak("that's all for today ..... have a good day")
        elif 'who are you' in query:
            speak("i am friday your virtual assistant")

        elif 'battery' in query:
            battery = psutil.sensors_battery()
            percent = str(battery.percent)
            speak(f"the battery percentage is{percent}percent")
        elif 'google search' in query:
            speak("what do you wanna search for sir ")
            a=takeCommand()
            speak("right away sir............. opening the results in google") 
            user_input=a
            google_search=requests.get('https://google.com/search?q='+user_input)
            soup=bs4.BeautifulSoup(google_search.text, 'html.parser')
            search_results=soup.select( r'a' )
            for link in search_results[2:3]:
                actual_link=link.get('href')
                print(actual_link)
                webbrowser.open('https://google.com/'+actual_link)
        elif 'how old are you':
            speak("well sir i'm younger than you but much wiser")
                   
        
        time.sleep(5)
