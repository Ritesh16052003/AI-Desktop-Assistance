import speech_recognition as sr
import pyttsx3
import wikipedia
import webbrowser
import os
import datetime
import subprocess
import ctypes
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL

# Initialize the speech engine for text-to-speech
engine = pyttsx3.init()

# Set voice property (0 for male, 1 for female)
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)  # Change to 'voices[0].id' for male voice

# Set speech rate (default is 200, you can modify it)
engine.setProperty('rate', 180)

# Volume control setup using pycaw (Python Core Audio Windows bindings)
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

# Path to WebDriver (update this with your actual WebDriver path)
chrome_driver_path = "C:\\path\\to\\chromedriver.exe"  # Replace with your ChromeDriver path

def speak(audio):
    """Convert text to speech."""
    engine.say(audio)
    engine.runAndWait()

def wish_me():
    """Greet the user based on the time of the day."""
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning!")
    elif hour >= 12 and hour < 18:
        speak("Good Afternoon!")
    else:
        speak("Good Evening!")
    speak("I am your desktop assistant. How may I assist you today?")

def take_command():
    """Take voice command from the user and return the recognized text."""
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        recognizer.pause_threshold = 1
        audio = recognizer.listen(source)

    try:
        print("Recognizing...")
        query = recognizer.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")
    except Exception as e:
        print("Sorry, I didn't catch that. Please say that again.")
        return "None"
    return query.lower()

# Volume control functions
def set_volume(level):
    """Set system volume to a specific level (0 to 100)."""
    volume_level = level / 100
    volume.SetMasterVolumeLevelScalar(volume_level, None)
    speak(f"Volume set to {level} percent")

def mute_volume():
    """Mute the system volume."""
    volume.SetMute(1, None)
    speak("System volume muted")

def unmute_volume():
    """Unmute the system volume."""
    volume.SetMute(0, None)
    speak("System volume unmuted")

# System control functions
def shutdown():
    """Shut down the system."""
    speak("Shutting down the system.")
    subprocess.call('shutdown /s /t 1')

def restart():
    """Restart the system."""
    speak("Restarting the system.")
    subprocess.call('shutdown /r /t 1')

def sleep():
    """Put the system to sleep."""
    speak("Putting the system to sleep.")
    ctypes.windll.powrprof.SetSuspendState(0, 1, 0)

# Browser Automation Functions
def open_browser():
    """Open a browser using Selenium."""
    speak("Opening Chrome browser.")
    driver = webdriver.Chrome(executable_path=chrome_driver_path)
    return driver

def search_google(query):
    """Search for something on Google."""
    speak(f"Searching Google for {query}")
    driver = open_browser()
    driver.get("https://www.google.com")
    search_box = driver.find_element_by_name("q")
    search_box.send_keys(query)
    search_box.send_keys(Keys.RETURN)


def search_youtube(query):
    """Search for a term on YouTube."""
    speak(f"Searching YouTube for {query}")
    driver = open_browser()
    driver.get("https://www.youtube.com")

    # Wait for the page to load completely
    time.sleep(2)

    # Find the search box element
    search_box = driver.find_element_by_name("search_query")
    search_box.send_keys(query)
    search_box.send_keys(Keys.RETURN)

    # Optionally, you can wait for results to load
    time.sleep(2)


def perform_task(query):
    """Execute the task based on user command."""
    if 'wikipedia' in query:
        speak('Searching Wikipedia...')
        query = query.replace("wikipedia", "")
        try:
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            speak(results)
        except wikipedia.exceptions.DisambiguationError as e:
            speak("There are multiple results for this term. Please be more specific.")
        except wikipedia.exceptions.PageError:
            speak("Sorry, I couldn't find anything on that topic.")
    elif 'search google' in query:
        search_term = query.replace("search google for", "").strip()
        search_google(search_term)
    elif 'open youtube' in query:
        speak("Opening YouTube.")
        webbrowser.open("https://www.youtube.com")
    elif 'open google' in query:
        speak("Opening Google.")
        webbrowser.open("https://www.google.com")
    elif 'open stack overflow' in query:
        speak("Opening Stack Overflow.")
        webbrowser.open("https://www.stackoverflow.com")
    elif 'volume' in query:
        if 'mute' in query:
            mute_volume()
        elif 'unmute' in query:
            unmute_volume()
        else:
            try:
                volume_level = int(query.split(' ')[-1])
                set_volume(volume_level)
            except ValueError:
                speak("Sorry, I couldn't understand the volume level.")
    elif 'shutdown' in query:
        shutdown()
    elif 'restart' in query:
        restart()
    elif 'sleep' in query:
        sleep()
    elif 'time' in query:
        str_time = datetime.datetime.now().strftime("%H:%M:%S")
        speak(f"The time is {str_time}")
    elif 'exit' in query or 'stop' in query:
        speak("Goodbye!")
        exit()
    else:
        speak("Sorry, I don't know how to handle that request.")

if __name__ == "__main__":
    wish_me()  # Greet the user
    while True:
        query = take_command()  # Listen to the user's command
        if query != "None":
            perform_task(query)  # Perform the task based on the command
