import streamlit as st
import win32com.client
import speech_recognition as sr
import webbrowser
import pyautogui
import wikipedia
import pygame

# Initialize Text-to-Speech engine
speaker = win32com.client.Dispatch("SAPI.SpVoice")

# Function to speak text
def say(text):
    speaker.Speak(text)

# Function to recognize voice command
def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        col1, col2 = st.columns([1, 10])  # Adjust the width ratio as needed
        with col1:
            st.image("D:\Ao\CS\\5 Python\Code\DScProjects\\begginer-guide\\ai-assistant.png")
        with col2:
            st.write("Listening...", text_align="center")  # Align text to the center
        
        try:
            audio = r.listen(source,0,3)
            st.write("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            col3, col4 = st.columns([1,10])
            with col3:
                st.image("D:\Ao\CS\\5 Python\Code\DScProjects\\begginer-guide\man.png")
            with col4: 
                st.write("You said: ")
            return query
        except sr.UnknownValueError:
            st.write("Sorry, I didn't catch that. Could you please repeat?")
            say("Could you please repeat that?")
            return ""
        except sr.RequestError as e:
            st.write(f"Could not request results from Google Web Speech Recognition service; {e}")
            return ""

# Function to play music
def play_music(file_path):
    pygame.mixer.init()
    pygame.mixer.music.load(file_path)
    pygame.mixer.music.play()

# Main function
def main():
    st.title("Voice-Controlled Assistant")

    # Main program loop
    while True:
        query = takeCommand().lower()
        st.write(query)

        if "open youtube" in query:
            say("Opening YouTube sir...")
            webbrowser.register('chrome', None, webbrowser.BackgroundBrowser("C:/Program Files/Google/Chrome/Application/chrome.exe"))
            webbrowser.get('chrome').open("https://www.youtube.com")
        
        elif "search wikipedia for" in query:
            topic = query.lower().replace("search wikipedia for", "").strip()
            say(f"Searching Wikipedia for {topic}...")
            topic = topic.replace(" ", "_")  # Replace spaces with underscores for Wikipedia search URL
            webbrowser.open(f"https://en.wikipedia.org/wiki/{topic}",1)
            # try:
            #     summary = wikipedia.summary(topic, sentences=1)
            #     st.write(summary)
            #     say(summary)

            # except wikipedia.exceptions.DisambiguationError as e:
            #     say(f"Disambiguation Error: {e.options}")
            # except wikipedia.exceptions.PageError:
            #     say("Page not found on Wikipedia.")

            try:
                summary = wikipedia.summary(topic, sentences=1)
                st.write(summary)
                say(summary)
            except wikipedia.exceptions.DisambiguationError as e:
                say(f"Disambiguation Error: {e.options}")
            except wikipedia.exceptions.PageError:
                say("Page not found on Wikipedia.")
            except wikipedia.exceptions.WikipediaException as e:
                say(f"Wikipedia Error: {str(e)}")


        
        elif "play" in query and "on youtube" in query:
            song = query.lower().replace("play", "").replace("on youtube", "").strip()
            say(f"Searching YouTube for {song}...")
            song = song.replace(" ", "+")  # Replace spaces with '+' for YouTube search URL
            webbrowser.open(f"https://www.youtube.com/results?search_query={song}", 1) # 1 for incognito
        
        elif "mute" in query:
            say("Got it")
            pyautogui.press('m')
        elif "unmute" in query:
            say("Got it")
            pyautogui.press('m')
        elif "play" in query:
            say("Got it")
            pyautogui.press('k')
        elif "play" in query:
            say("Got it")
            pyautogui.press("k")
        
        elif "close tab"  in query:
            say("Closing the current tab...")
            pyautogui.hotkey("ctrl","w")
        
        elif "finally sleep" in query.lower():
            say("Going to sleep sir")
            pyautogui.hotkey("ctrl","shift","w")
            exit()
        
        elif "close all tabs" in query:
            say("Closing all the tabs")
            pyautogui.hotkey("ctrl","shift","w")
        elif "your name" in query:
            say("I am Jarvis")
            st.write("I'm Jarvis")
        elif "hello" in query:
            say("Hello! How can I help you today?")
            st.write("Hello! How can I help you today?")
        elif "How are you" in query or "how r u" in query:
            say("Perfect sir")
            st.write("Perfect Sir")
        elif "remember that" in query:
            with open("remember.txt", "w") as f:  # Open the file in text mode ("w")
                query = query.replace("remember that", "")
                f.write(query)  # Writing the string directly
                say("remembered")
        elif "what do you remember" in query:
             with open("remember.txt","r") as f1:
                text = f1.read()
                say(text)
                st.write(text)
        elif "clear history" in query:
            say("Okay sir")
            pyautogui.hotkey("ctrl","r")
        
if __name__ == "__main__":
    main()
