import win32com.client
import speech_recognition as sr
import webbrowser
import pyautogui
import pygame
import time
import wikipedia

speaker = win32com.client.Dispatch("SAPI.SpVoice")

def say(text):
    speaker.Speak(text)

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        print("Listening...")
        try:
            audio = r.listen(source)
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User Said: {query}")
            return query
        except sr.UnknownValueError:
            print("Sorry, I didn't catch that. Could you please repeat?")
            return ""
        except sr.RequestError as e:
            print(f"Could not request results from Google Web Speech Recognition service; {e}")
            return ""

def play_music(file_path):
    pygame.mixer.init()
    pygame.mixer.music.load(file_path)
    pygame.mixer.music.play()


if __name__=='__main__':
    print("Running!!!")
    while True:
        query = takeCommand()
        print(query)
        if "open youtube" in query.lower():
            say("Opening YouTube sir...")
            webbrowser.register('chrome', None, webbrowser.BackgroundBrowser("C:/Program Files/Google/Chrome/Application/chrome.exe"))
            webbrowser.get('chrome').open("https://www.youtube.com")
        elif "search wikipedia for" in query.lower():
            topic = query.lower().replace("search wikipedia for", "").strip()
            say(f"Searching Wikipedia for {topic}...")
            topic = topic.replace(" ", "_")  # Replace spaces with underscores for Wikipedia search URL
            webbrowser.open(f"https://en.wikipedia.org/wiki/{topic}",1)
            try:
                summary = wikipedia.summary(topic, sentences=2)
                say(summary)
            except wikipedia.exceptions.DisambiguationError as e:
                say(f"Disambiguation Error: {e.options}")
            except wikipedia.exceptions.PageError:
                say("Page not found on Wikipedia.")
        elif "play" in query.lower() and "on youtube" in query.lower():
            song = query.lower().replace("play", "").replace("on youtube", "").strip()
            say(f"Searching YouTube for {song}...")
            song = song.replace(" ", "+")  # Replace spaces with '+' for YouTube search URL
            webbrowser.open(f"https://www.youtube.com/results?search_query={song}", 1) # 1 for incoginito
        elif "mute" in query.lower():
            say("Got it")
            pyautogui.press('m')
        elif "play" in query.lower():
            say("Got it")
            pyautogui.press("k")
        elif "close tab"  in query.lower():
            say("Closing the current tab...")
            pyautogui.hotkey("ctrl","w")
        elif "finally sleep" in query.lower():
             say("Going to sleep sir")
             exit()
        elif "close all tabs" in query:
            say("Closing all the tabs")
            pyautogui.hotkey("ctrl","shift","w")
        elif "open new tab" in query:
            say("opening new tab sir")
            pyautogui.hotkey("ctrl","t")
        elif "open new incognito tab" in query:
            say("opening new incoginito tab")
            pyautogui.hotkey("ctrl","shift","n")