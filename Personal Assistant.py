import webbrowser
import speech_recognition
import pyttsx3
import subprocess
import datetime
import os
import os.path
import datetime as dt

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from requests import Request

SCOPES = [" ---- paste your google api code here ----"]

GREEN_TEXT = '\033[91m'
RED_TEXT = '\033[92m'
RESET_TEXT_COLOR = '\033[0m'

# Initialize the text-to-speech engine
engine = pyttsx3.init()
recognizer = speech_recognition.Recognizer()

def speak(text):
    engine.say(text)
    engine.runAndWait()

# Function to get the appropriate greeting based on the current time
def get_greeting():
    current_hour = datetime.datetime.now().hour
    if 5 <= current_hour < 12:
        return "Good morning"
    elif 12 <= current_hour < 17:
        return "Good afternoon"
    else:
        return "Good evening"

def listen_for_trigger_word():
    with speech_recognition.Microphone() as mic:
        print(GREEN_TEXT+"Listening for the trigger word 'jarvis'..."+RESET_TEXT_COLOR)
        while True:
            audio = recognizer.listen(mic)
            try:
                text = recognizer.recognize_google(audio)
                
                if "jarvis" in text.lower():
                    speak("Hello Sir")
                    return

            except speech_recognition.UnknownValueError:
                pass

            except speech_recognition.RequestError as e:
                print("Error occurred while requesting results; {0}".format(e))    

def googlecalendar():
    creds = None

    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json")

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            
            # Edit this section
            
            flow = InstalledAppFlow.from_client_secrets_file("----- dowload the credentials file from google cloud and paste here -----", SCOPES)
            creds = flow.run_local_server(port=0)

        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("calendar", "v3", credentials=creds)

        event = {
            "summary": "My Python Event",
            "location": "Somewhere",
            "description": "Some more details on this awesome event",
            "colorId": 6,
            "start": {
                "dateTime": "2023-10-25T19:00:00+05:30",
                "timeZone": "Asia/Kolkata"
            },

            "end": {
                "dateTime": "2023-10-25T20:00:00+05:30",
                "timeZone": "Asia/Kolkata"
            },

            "recurrence": [
                "RRULE:FREQ=DAILY;COUNT=3"
            ],
            "attendees": [
                
                # Edit this section
                
                {"email": "<  ----  paste your email address here  ----  >"},
                # You can add the attendee emails here
                {"email": ""}
            ]
        }

        event = service.events().insert(calendarId="primary", body=event).execute()

        print(f"Event created: {event.get('htmlLink')}")

    except HttpError as error:
        print("An error occurred:", error)

def start_recognition():
    with speech_recognition.Microphone() as mic:
        speak("How may I be at your service .")
        
        while True:
            recognizer.adjust_for_ambient_noise(mic,duration=0.2)
            audio = recognizer.listen(mic)
            
            try:
                text = recognizer.recognize_google(audio)
                text = text.lower()
                print(GREEN_TEXT+f"Detected : "+RESET_TEXT_COLOR,RED_TEXT+text+RESET_TEXT_COLOR)
                speak("Command : " + text)

                if "close all" in text.lower():
                    str="Closing speech recognition."
                    print(GREEN_TEXT+str+RESET_TEXT_COLOR)
                    speak(str)
                    break

                elif "open chrome" in text.lower() :
                    str="Opening Google Chrome."
                    print(GREEN_TEXT+str+RESET_TEXT_COLOR)
                    speak(str)

                    # Edit this section
                    webbrowser.open(" -------- paste the link for google chrome from your file explorer --------  ") 
                
                elif "open firefox" in text.lower() :
                    str="Opening firefox"
                    print(GREEN_TEXT+str+RESET_TEXT_COLOR)
                    speak(str)
                    subprocess.Popen([" -------- paste the link for firefox from your file explorer --------  "])  
                
                elif "open personal mail" in text.lower() :
                    str="Opening personal mail."
                    print(GREEN_TEXT+str+RESET_TEXT_COLOR)
                    speak(str)
                    webbrowser.open(" -------- paste the link for personal mail id  --------  ")  
                
                elif "open windows terminal" in text.lower() :
                    str="Opening windows terminal."
                    print(GREEN_TEXT+str+RESET_TEXT_COLOR)
                    speak(str)
                    subprocess.Popen(["wt"]) 

                elif "open prime videos" in text.lower() :
                    str="Opening prime videos"
                    print(GREEN_TEXT+str+RESET_TEXT_COLOR)
                    speak(str)
                    webbrowser.open(" ------- paste your prime video login link here ------- ")     

                elif "sleep mode" in text.lower() :
                    str="sleep mode on"
                    print(GREEN_TEXT+str+RESET_TEXT_COLOR)
                    speak(str)
                    command = "rundll32.exe powrprof.dll,SetSuspendState 0,1,0"
                    os.system(command)

                elif "shutdown system" in text.lower() :
                    str="shuting down"
                    print(GREEN_TEXT+str+RESET_TEXT_COLOR)
                    speak(str)
                    command = "shutdown /s /f /t 0"
                    os.system(command)    
                
                elif "lock the laptop" in text.lower() :
                    str="locking"
                    print(GREEN_TEXT+str+RESET_TEXT_COLOR)
                    speak(str)
                    command = "rundll32.exe user32.dll,LockWorkStation"
                    os.system(command) 

                elif "add task to calendar" in text.lower() :
                    str="loading"
                    print(GREEN_TEXT+str+RESET_TEXT_COLOR)    
                    speak(str)
                    googlecalendar()

                # ....

            # You can add the processes here 
            
            except speech_recognition.UnknownValueError:
                print(GREEN_TEXT+"Could you please repeat it again sir"+RESET_TEXT_COLOR)
                speak("Could you please repeat it again.")

            except speech_recognition.RequestError as e:
                print("Error occurred while requesting results; {0}".format(e))    
                speak("Error occurred while requesting results.")


if __name__ == "__main__":
    greeting = get_greeting()
    speak(f"{greeting}. Your personal assistant is ready!")
    
  #  listen_for_trigger_word()
    start_recognition()  
