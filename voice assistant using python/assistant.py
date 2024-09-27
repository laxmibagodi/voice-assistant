import pywhatkit
import speech_recognition as sr
import pyttsx3
import datetime
import webbrowser
import wikipedia
import os
import subprocess
import wolframalpha
import json
import random
import operator
import winshell
import pyjokes
import feedparser
import ctypes
import time
import requests
import shutil
from twilio.rest import Client
from clint.textui import progress
import win32com.client as wincl
from urllib.request import urlopen

engine = pyttsx3.init()
# Female voice settings
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)  # Index 1 is usually a female voice

def greet():
    hour = datetime.datetime.now().hour
    if 0 <= hour < 12:
        return "Good morning!"
    elif 12 <= hour < 18:
        return "Good afternoon!"
    else:
        return "Good evening!"
    
def speak(text):
    engine.say(text)
    engine.runAndWait()

def listen():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source, timeout=15, phrase_time_limit=2 )         
            
    try:
        print("Recognizing...")
        query = recognizer.recognize_google(audio, language='en-in')
        print(f"You: {query}")
        return query.lower()
    except Exception as e:
        print(e)
        return ""

def execute_command(query):
    if 'open google' in query:
        print("Assistant: opening google, please wait!\n")
        speak("opening google, please wait!")
        webbrowser.open("https://google.com")
    elif 'wikipedia' in query:
        speak('Searching Wikipedia...')
        query = query.replace("wikipedia", "")
        results = wikipedia.summary(query, sentences = 3)
        speak("According to wikipedia...")
        print("results")
        speak(results)
    elif 'open youtube' in query or 'youtube' in query:
        print("Assistant: opening youtube, please wait!\n")
        speak("opening youtube, please wait!")
        webbrowser.open("https://youtube.com")
    elif 'nothing' in query:
        print("Assistant: ohh fine!\nis there anything i can do for you?\n")
        speak("ohhhh!,fine,is there anything i can do for you?")
    elif 'open linkedin' in query or 'linkedin' in query:
        print("Assistant: opening linkedin, please wait!\n")
        speak("opening linkedin, please wait!")
        
        webbrowser.open("https://linkedin.com")
    elif 'open wikipedia' in query or 'wikipedia' in query:
        print("Assistant: opening wikipedia, please wait!\n")
        speak("opening wikipedia, please wait!")
        webbrowser.open("https://wikipedia.com")
    
    elif 'time' in query:
        now = datetime.datetime.now().strftime("%H:%M")
        print(f"Assistant: The current time is {now}\n")
        speak(f"The current time is {now}")
        
    elif 'date' in query:
        today = datetime.datetime.now().strftime("%B %d, %Y")
        print(f"Assistant: Today's date is {today}\n")
        speak(f"Today's date is {today}")
    elif 'thank you' in query or 'thanks' in query or 'thank u' in query:
        print("Assistant: You're welcome!\n")
        speak("You're welcome!")

    elif 'what is your age' in query:
        print("Assistant: i am computer program dude,seriously u r asking me this?\n")
        speak("i am computer program dude,seriously u r asking me this?")
    elif 'your name' in query:
        print("Assistant: i am virtual voice assistant designed to assist people.\n")
        speak("i am virtual voice assistant designed to assist people.")
    elif 'who created you' in query or 'created' in query:
        print("Assistant: a project group of 2nd sem under guide - Dr.manikamma mam created me!\n the students of group are Laxmi,Nivedita,Priya,Mallamma,Sahana")
        speak("a project group of 2nd semester under guide doctor manikamma mam created me! the students of group are, laxmi, nivedita,priya,mallamma, sahana")
    elif 'hello' in query or 'hey' in query:
        print(" hello! how are you?\n How can i assist you today?\n")
        speak(" hello ! how are you? How can i assist you today?")
    elif 'funny' in query or 'not funny' in query or 'kidding' in query or 'just joking' in query:
        print("HaHaHaHaHa! \n is there anything i can do for you?")
        speak("HaHaHaHaHa! is there anything i can do for you?")
    elif 'how are you' in query or 'how about you' in query:
        print(" Always good" ) 
        speak(" Always good ,what can i do for you! ")


    elif 'i am fine' in query or 'good' in query or 'i am doing good' in query or 'fine' in query:
        speak("great to hear that! what can i do for u today?")

    elif 'take a note' in query:
            print("Sure! Write it here")
            speak("Sure! Write it here")
            def save_note():
                desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
            
                with open(os.path.join("C:\\Users\\DELL\\Downloads\\notes.txt", "note.txt"), "a") as f:
                    f.write("")
            os.startfile(os.path.join("C:\\Users\\DELL\\Downloads\\notes.txt", "note.txt"))
            save_note()
    elif  'where' in query or 'about' in query or 'how' in query or 'why' in query:
        information = query
        try:
            info = wikipedia.summary(information, sentences=2)
            print(info)
            speak(info)
            print("is there anything i can do for you?")
            speak("is there anything i can do for you?")
        except Exception as e:
            speak("sorry! i couldnt find anything related to query")


    elif 'exit' in query:
        print("Assistant: Byeee! \n it was nice talking to you, see you soon")
        speak("Byeee!,it was nice talking to you, see you soon")
        exit()

    elif "play" in query:
        song =query.replace('play', "")
        speak("playing" + song)
        pywhatkit.playonyt(song)

    elif "change name" in query:
        speak("What would you like to call me? ")
        assname = listen()
        speak("Thanks for naming me")
    elif 'joke' in query:
        joke = pyjokes.get_joke()
        print(joke)
        speak(joke)
    elif "calculate" in query:
        app_id = "3JALK5-EVPYG5YPHR"
        client = wolframalpha.Client(app_id)
        indx = query.lower().split().index('calculate') 
        query = query.split()[indx + 1:]
        res = client.query(' '.join(query))
        answer = next(res.results).text
        print("The answer is " + answer)
        speak("The answer is " + answer)
    elif 'search' in query or 'play' in query:
        query = query.replace("search", "")
        query = query.replace("play", "")
        webbrowser.open(query)
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
        speak("Recycle Bin Recycled")
    elif "restart" in query:
        subprocess.call(["shutdown", "/r"])
    elif "hibernate" in query or "sleep" in query:
        speak("Hibernating")
        subprocess.call("shutdown / h")
    elif "log off" in query or "sign out" in query:
        speak("Make sure all the application are closed before sign-out")
        time.sleep(5)
        subprocess.call(["shutdown", "/l"])
    elif "write a note" in query:
        speak("What should i write?")
        note = listen()
        file = open('assistant.txt', 'w')
        file.write(note)
        speak("note taken")
    elif "show note" in query:
       speak("Showing Notes")
       file = open("assistant.txt", "r")
       print(file.read())
       speak(file.read(6))

    elif "message" in query or "can you send me a message" in query :
        account_sid = 'AC370df999ff3e9bf6cc3e0a91c45616c7'
        auth_token = '6a9a83c7f351f291d55748fcde0a99ac'
        client = Client(account_sid, auth_token)
    
        speak("What should I send?")
        message_body = listen()
    
        message = client.messages.create(
            body=message_body,
            from_='+12566078744',
            to='+916362084301'
        )
        print("message sent successfully")
        speak("message sent successfully")
        print("id : ", message.sid)
    elif "Good Morning" in query:
        speak("A warm" +query)
        speak("How are you?")
        speak(assname)

    elif 'powerpoint presentation' in query:
        speak("opening Power Point presentation")
        power = r"C:\\Users\\DELL\\Downloads\\VIRTUAL VOICE ASSISTANT   USING PYTHON.pptx"
        os.startfile(power)
    elif "what is" in query or "who is" in query:
        client = wolframalpha.Client("3JALK5-T3WQ2X5QYH")
        res = client.query(query)
        try:
            print (next(res.results).text)
            speak (next(res.results).text)
        except:
            print ("No results")

    else:
        speak("I'm sorry, I didn't understand that.")



def run_assistant():
    speak(greet())
    speak("i am virtual voice assistant How can I assist you today?")
    print(greet())
    print("i am virtual voice assistant. How can I assist you today?")


    while True:
        query = listen()
        if query:
            execute_command(query)

if __name__ == "__main__":
   
    run_assistant()