
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import smtplib


def speak(str):
    from win32com.client import Dispatch
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(str)


def wishme():
    hour = int(datetime.datetime.now().hour)
    if 6 <= hour < 12:
        speak("good Morning")
    elif 12 <= hour < 18:
        speak("Good Afternoon")
    else:
        speak("good Evening")

    speak("I'am xenos sir. Please tell me how may i help you")


def take_command():
    r = sr.Recognizer()
    r.energy_threshold = 400
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source, duration=5)
        print("listening........")
        r.pause_threshold = 1
        audio = r.listen(source)
    try:
        print("Recognising.......")
        query = r.recognize_google(audio, language='en-in')
        print(f"user said : {query}\n")

    except Exception as e:

        print("say that again please.....")
        return "None"
    return query


def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('sahilgautam2k16@gmail.com', 'password')
    server.sendmail('sahilgautam2k16@gmail.com', to, content)
    server.close()


if __name__ == '__main__':
    print("Welcome to Xenos beta Test version 1.16")
    wishme()
    while True:
        query = take_command().lower()

        if 'wikipedia' in query:
            speak("searching on wikipedia")
            query = query.replace("wikipedia", "")
            result = wikipedia.summary(query, sentences=2)
            speak("accoding to wikipedia")
            print(result)
            speak(result)
        elif 'open youtube' in query:
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            webbrowser.open("google.com")

        elif 'open stackover flow' in query:
            webbrowser.open("stackoverflow.com")

        elif "the time" in query:
            str_time = datetime.datetime.strftime("H%:M%:S%")
            speak(f"sir the time is {str_time}")
        elif 'open code' in query:
            path = "C:\\Users\\Amang\\AppData\\Local\\aMicrosoft VS Code\\Code.exe"
            os.startfile(path)
        elif 'open game' in query:
            path = "C:\\Riot Games\\Riot Client\\RiotClientServices.exe"
            os.startfile(path)

        elif 'send email to speedy' in query:
            try:
                speak("what should i say")
                content = take_command()
                to = "speedyaashi@gmail.com"
                sendEmail(to, content)
                speak("email has been sent")
            except Exception as e:
                speak("sorry i'm failed to send the mail")

