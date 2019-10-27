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

def shutdown(confirmation):
    if confirmation == 'y':
        os.system("shutdown /s /t 1")
    else:
        return


def restart(confirmation):
    if confirmation == 'y':
        os.system("shutdown /r /t 1")
    else:
        return

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
# ---------------------------music---------------------------
        elif 'play music' in query:
            music_folder = "F:\\musical\\"
            music = ['rabta', 'barsaat']
            ok = random.choice(music)
            random_music = music_folder + ok + '.mp3'
            os.system(random_music)

            speak('Okay, here is your music! Enjoy!', 0)

        elif 'change music' in query:
            music_folder1 = "F:\\musical\\"
            music1 = ['rabta', 'barsaat']
            ok1 = random.choice(music1)
            while ok1 == ok:
                ok1 = random.choice(music1)
            random_music1 = music_folder + ok1 + '.mp3'
            os.system(random_music1)

            speak('Okay, here is your music! Enjoy!', 0)
            
# ---------------------------Shutdown and Restart---------------------------            
        elif'shutdown' in query:
                speak("Are you sure you want to shutdown?")
                confirmation = LISTEN()
                shutdown(confirmation[0])
           
        elif'restart' in query:
                speak("Are you sure you want to restart?")
                confirmation = LISTEN()
                restart(confirmation[0])

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
