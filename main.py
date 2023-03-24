import win32com.client
import time
import speech_recognition as sr

speaker = win32com.client.Dispatch("SAPI.SpVoice")


def tell(audio):
    speaker.Speak(audio)


def greetings():
    hour = int(time.strftime("%H"))
    if hour >= 5 and hour < 12:
        tell("Good Morning!!")
    elif hour >= 12 and hour <= 18:
        tell("Good Evening")
    else:
        tell("Good Night.")


def micInput():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("I am hearing...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language="en-in")
        tell(f"Speaker said: {query}")
    except Exception as e:

        print("Say that again!")
        return "None"
    return query


if __name__ == "__main__":
    greetings()
    micInput()
