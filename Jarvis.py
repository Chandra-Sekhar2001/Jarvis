import win32com.client
import webbrowser
import smtplib
import random
import speech_recognition as sr
import wikipedia
import datetime
import os
import sys

speaker = win32com.client.Dispatch("SAPI.SpVoice")



def speak(audio, a):
    if a == 1:
        print('Command on board: ' + audio)
    speaker.Speak(audio)


def wish():
    currenttime = int(datetime.datetime.now().hour)
    if currenttime >= 0 and currenttime < 12:
        speak('Good morning', 0)
    if currenttime >= 12 and currenttime < 18:
        speak('Good afternoon', 0)
    if currenttime >= 18 and currenttime < 24:
        speak('Good evening', 0)


def LISTEN():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Destination please ")
        audio = r.listen(source)
    try:
        text = r.recognize_google(audio, language="en-IN")
    except sr.UnknownValueError:
        text = LISTEN()
        print("Audio is unintelligible")
    except sr.RequestError as e:
        text = ""
        speak("Try again")
        print(f"Can't obtain results {e}")
    return text

if __name__ == '__main__':
    wish()
    speak('Hello Sir,I am your digital assistant Jarvis', 0)
    speak('How may I help you', 1)
    sites = {'youtube': 'https://www.youtube.com',
             'google': 'https://www.google.com', 'gmail': 'www.gmail.com'}
    while True:
        query = LISTEN().lower()
# ---------------------------browser based---------------------------

        if 'open ' in query:
            query = query.replace("open ", "")
            speak('okay', 0)
            webbrowser.open(sites[query])

# --------------------------- wishing---------------------------
        elif "what\'s up" in query or 'how are you' in query:
            stMsgs = ['Just doing my thing!', 'I am fine!',
                      'Nice!', 'I am nice and full of energy']
            speak(random.choice(stMsgs), 0)

        elif 'hello' in query:
            speak('Hello Sir', 0)

        elif 'bye' in query:
            speak('Bye Sir, have a good day.', 0)
            sys.exit()

# ---------------------------Searchiung word---------------------------
        else:
            query = query   # I THINK ITS NOT REQUIRED ONCE CHECK.
            speak('Searching...', 1)
            try:
                results = wikipedia.summary(query, sentences=2)
                speak('Got it.', 1)
                speak('WIKIPEDIA says - ', 1)
                speak(results, 1)
            except:
                webbrowser.open('www.google.com')

        speak('Next Command! Sir!', 1)
