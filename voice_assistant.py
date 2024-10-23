import pyttsx3 as P
import speech_recognition as sr
import pyaudio
import os
import subprocess
import requests
import webbrowser
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume # To control the volume output of the speaker on my device
from comtypes import CLSCTX_ALL  # To use the microphone
import pyautogui # To use the keyboard interface
import time # To use 

# Initialize the text-to-speech engine
engine = P.init()

rate = engine.getProperty('rate')
engine.setProperty('rate', 150)  # Speed of speech
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)  # Use the second voice (female on some systems)

def talk(text):
    """Speaks the text passed as argument"""
    engine.say(text)
    engine.runAndWait()

# Voice recognition
r = sr.Recognizer()

def listen():
    """Listens to user voice input and converts it to text"""
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source, 1.2)
        print("Listening...")
        talk(" You didn't give  any command. How can I help you?")
        audio = r.listen(source)
        try:
            text = r.recognize_google(audio)
            print(f'You said: {text}')
            return text.lower()
        except sr.UnknownValueError:
            talk("Sorry, I didn't catch that. Could you repeat please?")
            return listen()  # Recursively listen again
        except sr.RequestError as e:
            talk(f"Sorry, there was an issue with the request: {e}")
            return None

# Commands
import subprocess
import webbrowser

def open_application(app_name):
    """Opens the specified application or website."""
    app_name = app_name.lower()
    
    # Local applications
    if "notepad" in app_name:
        subprocess.Popen('notepad.exe')
        talk("Opening Notepad.")
    elif "calculator" in app_name:
        subprocess.Popen('calc.exe')
        talk("Opening Calculator.")
    elif "word" in app_name:
        subprocess.Popen('winword.exe')
        talk("Opening Word.")
    elif "excel" in app_name:
            subprocess.Popen('excel.exe')
            talk("Opening Excel.")
    elif "powerpoint" in app_name:
        subprocess.Popen('powerpnt.exe')
        talk("Opening PowerPoint.")
    elif "whatsapp" in app_name:
        subprocess.Popen(r"C:\Users\PC\AppData\Local\WhatsApp\WhatsApp.exe")
        talk("Opening WhatsApp.")
    elif "chrome" in app_name or "browser" in app_name:
        subprocess.Popen(r"C:\Program Files\Google\Chrome\Application\chrome.exe")
        talk("Opening Chrome.")
    
    # Browser-based platforms
    elif "youtube" in app_name:
        webbrowser.open("https://www.youtube.com")
        talk("Opening YouTube.")
    elif "tiktok" in app_name:
        webbrowser.open("https://www.tiktok.com")
        talk("Opening TikTok.")
    elif "x" in app_name or "twitter" in app_name:
        webbrowser.open("https://www.twitter.com")
        talk("Opening X, formerly known as Twitter.")
    elif "wikipedia" in app_name:
        webbrowser.open("https://www.wikipedia.org")
        talk("Opening Wikipedia.")
    elif "trends in tiktok" in app_name:
        webbrowser.open("https://www.tiktok.com/trending")
        talk("Opening TikTok trending page.")
    elif "trends in twitter" in app_name:
        webbrowser.open("https://twitter.com/trending")
        talk("Opening Twitter trending page.")
    elif "trends in youtube" in app_name:
        webbrowser.open("https://www.youtube.com/feed/trending")
        talk("Opening YouTube trending page.")
    
    # If app is not supported
    else:
        talk("Sorry, I cannot open that application.")


def change_volume(action):
    """Increases or decreases the systems volume."""
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        current_volume = volume.GetMasterVolume()
        
        if action == "increase":
            new_volume = min(current_volume + 0.1, 1.0)  # Increase volume by 10%
        elif action == "decrease":
            new_volume = max(current_volume - 0.1, 0.0)  # Decrease volume by 10%

        volume.SetMasterVolume(new_volume, None)
        talk(f"Volume {'increased' if action == 'increase' else 'decreased'}.")
def set_volume(level):
    """Sets the system volume to the given percentage (0 to 100)"""
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        new_volume = min(max(level / 100, 0.0), 1.0)  # Ensure it's between 0.0 and 1.0
        volume.SetMasterVolume(new_volume, None)
        talk(f"Volume set to {level} percent.")


# # Usage example
# change_volume("increase")  # To increase volume
# change_volume("decrease")  # To decrease volume


def close_application(app_name):
    """Closes the specified application"""
    if "notepad" in app_name:
        os.system("taskkill /f /im notepad.exe")
        talk("Closing Notepad.")
    elif "chrome" in app_name or "browser" in app_name:
        os.system("taskkill /f /im chrome.exe")
        talk("Closing Chrome.")
    elif "calculator" in app_name:
        os.system("taskkill /f /im calc.exe")

    else:
        talk("Sorry, I cannot close that application.")

def reopen_recently_closed_tab(browser_name):
    if browser_name == "chrome" or browser_name == "firefox" or browser_name == "safari" or browser_name == "brave" or browser_name == "edge" :
        # Pressing Ctrl + Shift + T to reopen last closed tab
        pyautogui.hotkey('ctrl', 'shift', 't')
        talk(f"Reopening recently closed tab in {browser_name}.")
    else:
        talk(f"Sorry, reopening tabs recent {browser_name} is not supported yet or the application is not downloaded yet. can l download it for you now ")

# Example usage
# reopen_recently_closed_tab("chrome")


def shutdown_pc():
    """Shuts down the pc"""
    talk("Shutting down the computer.")
    os.system("shutdown /s /t 1")

def restart_pc():
    """Restarts the pc"""
    talk("Restarting the computer.")
    os.system("shutdown /r /t 1")

def search_web(query):
    """Searches the web for the given query using Google"""
    talk(f"Searching the web for {query}")
    webbrowser.open(f"https://www.google.com/search?q={query}")

def read_news():
    """Fetches and reads out the latest news headlines"""
    talk("Fetching the latest news.")
    api_key = "YOUR_NEWS_API_KEY"  # You need to get a News API key from https://newsapi.org
    url = f"https://newsapi.org/v2/top-headlines?country=us&apiKey={api_key}"
    response = requests.get(url)
    news_data = response.json()

    if news_data["status"] == "ok":
        talk("Here are the latest news headlines.")
        for i, article in enumerate(news_data['articles'][:5], start=1):
            talk(f"News {i}: {article['title']}")
    else:
        talk("Sorry, I could not fetch the news.")
        

# Main control loop
def voice_assistant():
    """Main Voice assistant loop to listen and perform actions"""
    talk("Hello! My name is Alice . How can l be of help to you today?")
    
    while True:
        command = listen()

        if command:
            if "open" in command:
                app_name = command.replace("open", "").strip()
                open_application(app_name)

            elif "close" in command:
                app_name = command.replace("close", "").strip()
                close_application(app_name)

            elif "shutdown" in command:
                shutdown_pc()
                break  # Exit after shutdown command

            elif "restart" in command:
                restart_pc()
                break  # Exit after restart command

            elif "search" in command:
                query = command.replace("search", "").strip()
                search_web(query)

            elif "news" in command:
                read_news()

            elif "who is your technician" in command or "who is your technical support" in command or "who is your tech support" in command or "who is your tech support person" in command or "who is your technical support  team" in command or "who do you work for" in command or "who created you " in command:
                talk("I am Alice. Your voice assistant. And my technician is 'duckworth-pixel'.")
             
            elif "change volume" in command or "change volume up" in command or "change volume down" in command or "increase volume" in command or "decrease volume" in command or "increase volume up" in command or "decrease volume down" in command or "volume up" in command or "volume down" in command or "volume up to 25" in command or "volume down to 0" in command or "volume up to 50" in command or "volume up to 75" in command or "volume up to 100" in command or "volume down to 25" in command or "volume down to 50" in command or "volume down to 75" in command or "volume down to 15" in command:
                change_volume(command)

            elif "stop" in command or "exit" in command or "terminate" in command or "goodbye" in command or "bye" in command or "see you" in command or "see you later" in command or "see you soon" in command:
                talk("Feel free to reach out if you'd like to be helped in any way  am always here for you! Bye for now")
                break

            else:
                talk("I'm not sure how to handle that. Please consult with my technicial for help.")

# Start the voice assistant
if __name__ == "__main__":
    voice_assistant()
