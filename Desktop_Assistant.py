
import smtplib as sm
import speech_recognition as sr
import datetime as dt
import webbrowser as wb
import pyaudio as py
import wikipedia as wk
import pyttsx3 as py #python library that will help us to convert text to speech.
import os
from playsound import playsound


engine = py.init("sapi5")  #Microsoft Speech API (SAPI5) is the technology for voice recognition and synthesis provided by Microsoft.
voices = engine.getProperty("voices") #getting details of current voices
# print(voices)
engine.setProperty("voices",voices[0])


def speak(audio):
    engine.say(audio)
    engine.runAndWait() #Without this command, speech will not be audible to us.


def wishme():
    
    hour=dt.datetime.now().hour
    
    if hour>=0 and hour<12:
        speak("good morning buddy")
    
    elif hour>=12 and hour<18:
        speak("good afternoon buddy")
    
    else:
        speak("good evening buddy")
    
    speak("I am your desktop assistant, Please tell me how may I help you")

def takeCommand():  #It takes microphone input from the user and returns string output
    
    r = sr.Recognizer()
    with sr.Microphone(1) as source:
        print("Listening...")
        r.pause_threshold = 0.8
        r.adjust_for_ambient_noise(source)
        audio = r.listen(source)
    
    try:
        print("Recognizing...")
        query=r.recognize_google(audio, language="en-in") #Using google API for speech recognition (speech to text)
        print(f"User said: {query}")
    
    except Exception as e:
        #print(e)
        # speak("Say that again please")
        print("Say that again please...")
        return "None"
    
    return query
        


def sendMail(to,content):
    
    server = sm.SMTP("smtp.gmail.com",587)
    server.starttls() #why is it used ?
    server.ehlo() #why is it used ?     When the STARTTLS command is used, the EHLO command must also be used.
    server.login("<mail-id-here>", "<password-here>")
    server.sendmail("<mail-id-here>", to, content)
    server.close()

if __name__=="__main__":
    
    # wishme()
    

    while True:
        
        query = takeCommand().lower()    
        
        if "wikipedia" in query:
            speak("Searching in wikipedia")
            query=query.replace("wikipedia", "")
            result = wk.summary(query, sentences = 2)
            speak("According to wikipedia")
            print(result)
            speak(result)

        elif "play music" in query:
            music_dir = r"C:\Users\User\Downloads\Songs" 
            songs = os.listdir(music_dir) #list all the songs present in the directory
            print(songs)
            os.startfile(os.path.join(music_dir,songs[0]))



        elif "open notepad" in query:
            notepad = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Accessories\Notepad"
            speak("Opening notepad for you")
            os.startfile(notepad)
        
        elif "open command prompt" in query:
            cmd = r"C:\Users\User\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\System Tools\Command Prompt"
            speak("Opening command prompt for you")
            os.startfile(cmd)

        elif "open youtube" in query:
            speak("opening youtube for you")
            wb.open("https://www.youtube.com")
        
        elif "open google" in query:
            speak("Opening google for you")
            wb.open("https://www.google.com")
        
        elif 'open stackoverflow' in query:
            speak("Opening stackoverflow for you")
            wb.open("https://stackoverflow.com")  
        
        elif "what is the time" in query:
            Time = dt.datetime.now().strftime("%H:%M:%S")
            speak(f"hey buddy, the current time is {Time}")
        
        elif "open vs code" in query:
            vsCode=r"C:\Users\User\AppData\Local\Programs\Microsoft VS Code\Code.exe"
            speak("opening vs code for you")
            os.startfile(vsCode)

        elif "send mail" in query:
            try:
                speak("Please tell me the content of the mail...")
                content = takeCommand()
                # # speak("to whom you want to send this mail")
                # to = takeCommand()
                # print(to)
                to = "<mail-id-here>"
                sendMail(to,content)
                speak("Mail has been sent")
            except Exception as e:
                #print(e)
                speak("sorry, Mail has not been sent")
        
        elif "bye thank you" in query:
            speak("Thank you buddy, I am happy to help you. Have a good day ahead")
            exit() 


