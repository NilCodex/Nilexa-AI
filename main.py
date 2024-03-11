import speech_recognition as sr

import win32com.client as wincom
import webbrowser


def say(text):
    speak = wincom.Dispatch("SAPI.SpVoice")
    text = f"{text}"
    speak.Speak(text)


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 0.5
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_sphinx(audio)
            print(f"user-said:{query}")
            return query
        except Exception as e:
            # exit("Can You please repeat it again :")
            return "I did not understand that, please repeat it again..."


if __name__ == '__main__':
    print('Welcome to Nilexa AI')
    say("Activate Nilexa")
    while True:

        print("listening...")
        query = takeCommand()
        say(query)
        sites = [["Youtube", "https://www.youtube.com/"], ["Instagram", "https://www.instagram.com/"],
                 ["Facebook", "https://www.facebook.com/"], ["Google", "https://www.google.com/"],
                 ["linkedin", "https://www.linkedin.in/"]]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                say(f"Opening {site[0]} ,Sir")
                webbrowser.open(site[1])


