import speech_recognition
import pyttsx3
import pywhatkit
import wikipedia
import webbrowser
def takeCommand():
    r = speech_recognition.Recognizer()
    with speech_recognition.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        r.energy_threshold = 250
        audio = r.listen(source, 0, 3)

        try: 
            print("Understanding...")
            query = r.recognize_azure(audio, language='en-id')
            print(f"You sain {query}\n")
        except Exception as e:
            print("Say that again")
            return "None" 
        return query
# Above in all files
# below code for engine
     
query = takeCommand().lower()
engine = pyttsx3.init("sapi5")
voices = engine.getProperty("voices")
engine.setProperty("voice", voices[0].id) 
engine.setProperty("rate",170)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def searchGoogle(query):
    if "google" in query:
        import wikipedia as googleScrap
        query = query.replace("Jarvis", "")
        query = query.replace("google search","")
        query = query.replace("google","")
        speak("This is what I found on google")

        try: 
            pywhatkit.search(query)
            reuslt = googleScrap.summary(query,1)
            speak(result)
        except:
            speak("No speakable output available")

def searchYoutube(query):
    if "youtube" in query:
        speak("This is what I found for your search!")
        query= query.replace("youtube search")
        query= query.replace("youtube","")
        query= query.replace("Jarvis","")
        web = "https://www.youtube.com/results?search_query="+ query
        webbrowser.open(web)
        pywhatkit.playonyt(query)
        speak("Done sir")

def searchWikipedia(query):
    if "wikipedia" in query:
        speak("Searching from wikipedia.")
        query= query.replace('wikipedia',"")
        query= query.replace("search wikipedia","")
        query= query.replace("Jarvis", "")
        results = wikipedia.summary(query, sentences=2)
        speak("According to wikipedia.")
        print(results)
        speak(results)