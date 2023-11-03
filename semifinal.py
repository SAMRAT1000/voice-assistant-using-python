import speech_recognition as sr
import pyttsx3
import datetime
import wikipedia
# import webbrowser

import imdb
import os
import win32com.client as wincl
import time
import subprocess

# from ecapture import ecapture as ec
import wolframalpha
import json
import requests

# import pyjokes
import random
import subprocess
# from urllib.request import urlopen
import smtplib

# import winshell
import ctypes
import smtplib
from email.message import EmailMessage
import pyttsx3


print("Loading your AI personal assistant - Jarvis One")

engine = pyttsx3.init("sapi5")
voices = engine.getProperty("voices")
engine.setProperty("voice", voices[0].id)


def speak(text):
    engine.say(text)
    engine.runAndWait()


def wishMe():
    hour = datetime.datetime.now().hour
    if hour >= 0 and hour < 12:
        speak("nuhmuhstay")

        speak("Good Morning")
        print("Nuhmustay Good Morning")
    elif hour >= 12 and hour < 18:
        speak("nuhmuhstay")

        speak("Good After noon")
        print("Nuhmustay Good Afternoon")
    else:
        speak("nuhmuhstay")
        # speak("  ")
        speak("Good evening")
        print("Nuhmustay,Good Evening")


def search_movie():
    # gathering information from IMDb
    moviesdb = imdb.IMDb()

    # search for title
    text = takeCommand()

    # passing input for searching movie
    movies = moviesdb.search_movie(text)

    speak("Searching for " + text)
    if len(movies) == 0:
        speak("No result found")
    else:
        speak("I found these:")

        for movie in movies:
            title = movie["title"]
            year = movie["year"]
            # speaking title with releasing year
            speak(f"{title}-{year}")

            info = movie.getID()
            movie = moviesdb.get_movie(info)

            title = movie["title"]
            year = movie["year"]
            rating = movie["rating"]
            plot = movie["plot outline"]

            # the below if-else is for past and future release
            if year < int(datetime.datetime.now().strftime("%Y")):
                speak(
                    f"{title}was released in {year} has IMDB rating of {rating}.\
                    The plot summary of movie is{plot}"
                )
                print(
                    f"{title}was released in {year} has IMDB rating of {rating}.\
                    The plot summary of movie is{plot}"
                )
                break

            else:
                speak(
                    f"{title}will release in {year} has IMDB rating of {rating}.\
                    The plot summary of movie is{plot}"
                )
                print(
                    f"{title}will release in {year} has IMDB rating of {rating}.\
                    The plot summary of movie is{plot}"
                )
                break


def takeCommand():
    """It takes microphone input from the user and returns string output"""
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("==============")
        print("   LISTNING  ")
        print("==============\n")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        statement = r.recognize_google(audio, language="en-in")
        print(f"User said:,{statement}\n")

    except Exception as e:
        # print(e)

        print("say that again please...")
        return "None"
    return statement.lower()


assistant_name = "jarvis"

wishMe()
speak(" ")
speak("I am your Assistant")
speak(assistant_name)
print(assistant_name)

print("\n")


def send_mail(receiver, subject, body):
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login("dumdumtrex@gmail.com", "fzlyahkedorkrqof")
    server.sendmail("dumdumtrex@gmail.com", "freestylewinds777@gmail.com", "heyelloo")
    email = EmailMessage()
    email["from"] = "dumdumtrex@gmail.com"
    email["To"] = receiver
    email["Subject"] = subject

    email.set_content(body)
    server.send_message(email)


def main_trex():
    speak("hello sir")
    speak("whome should i send the mail?")
    name = takeCommand()
    receiver = dict[name]
    speak("Speak the subject sir")
    subject = takeCommand()
    speak("Speak the message sir")
    body = takeCommand()
    send_mail(receiver, subject, body)
    print("Your email has been sent")


dict = {
    "t1": "freestylewinds777@gmail.com",
    "t2": "trex2333777@gmail.com",
    "t3": "trex1333777@gmail.com",
    
}

if __name__ == "__main__":
    while True:
        statement = takeCommand().lower()
        if statement == 0:
            continue

        elif "introduce yourself" in statement:
            speak("yes master i m " + str(assistant_name))

            speak("I am an Artificial Intelligence based program")
            speak(
                ". To be more precise,i am a smart computer program that understand human language through text or particular voice commands. Consequently i perform tasks for you master."
            )

        elif "send mail" in statement:
            main_trex()

        elif "wikipedia" in statement:
            speak("Searching Wikipedia...")
            statement = statement.replace("wikipedia", "")
            results = wikipedia.summary(statement, sentences=3)
            speak("According to Wikipedia")
            print(results)
            speak(results)
        elif "find me a movie" in statement:
            speak("master please tell movie name")
            search_movie()
        elif "open youtube" in statement:
            webbrowser.open_new_tab("https://www.youtube.com")
            speak("youtube is open now")
            time.sleep(5)

        elif "open google" in statement:
            webbrowser.open_new_tab("https://www.google.com")
            speak("Google chrome is open now")
            time.sleep(5)

        elif "open gmail" in statement:
            webbrowser.open_new_tab("gmail.com")
            speak("Google Mail open now")
            time.sleep(5)

        elif "time" in statement:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"the time is {strTime}")

        elif "who are you" in statement or "what can you do" in statement:
            speak(
                "I am"
                + str(assistant_name)
                + " version 1 point O your persoanl assistant. I am programmed to do minor tasks like"
                "opening youtube,google chrome,gmail and stackoverflow ,predict time,take a photo,search wikipedia,predict weather"
                "in different cities , get top headline news from times of india and you can ask me computational or geographical questions too!"
            )

        elif (
            "who made you" in statement
            or "who created you" in statement
            or "who discovered you" in statement
        ):
            speak("I was built by team dagger")
            print("I was built byteam dagger")

        elif "open stackoverflow" in statement:
            webbrowser.open_new_tab("https://stackoverflow.com/login")
            speak("Here is stackoverflow")

        elif "news" in statement:
            news = webbrowser.open_new_tab(
                "https://timesofindia.indiatimes.com/home/headlines"
            )
            speak("Here are some headlines from the Times of India,Happy reading")
            time.sleep(6)

        # elif "camera" in statement or "take a photo" in statement:
        #     ec.capture(0, "robo camera", "img.jpg")

        elif "open spotify" in statement:
            webbrowser.open("spotify.com")
            speak("Here you go with spotify")

        elif "open ppt" in statement:
            speak("opening Power Point presentation")

            power = r"C:\\J-A-R-V-I-S Sem 1 dt project\\final jarvis.pptx"
            # power = r"C:\\Users\\hp\\OneDrive\\Documents\\thanos\\jarvisppt.pptx"

            os.startfile(power)
        elif "the time" in statement:
            strTime = datetime.datetime.now().strftime("% H:% M:% S")
            speak(f"Sir, the time is {strTime}")

        elif "open code" in statement:
            codePath = (
                "C:\\Users\\hp\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            )
            os.startfile(codePath)

        elif "how are you" in statement:
            speak("I am fine, Thank you")
            speak("How are you, Sir")

        elif "fine" in statement or "good" in statement:
            speak("It's good to know that your fine")

        elif "is love" in statement:
            speak("It is 7th sense that destroy all other senses")

        elif "will you be my gf" in statement or "will you be my bf" in statement:
            speak("I'm not sure about, may be you should give me some time")

        elif "change name" in statement:
            speak("What would you like to call me, Sir ")
            assistant_name = takeCommand()
            speak("Thanks for naming me")
            speak(assistant_name)

        elif "what's your name" in statement or "What is your name" in statement:
            speak("My friends call me")
            speak(assistant_name)
            print("My friends call me", assistant_name)
        elif "exit" in statement:
            speak("Thanks for giving me your time")
            exit()

       
        elif (
            " was this suppose to be funnny " in statement
            or "i don't find it funny" in statement
        ):
            speak("Sorry i will try harder next time.")
        elif "why you came to world" in statement:
            speak("Thanks to team dagger . further It's a secret")
        elif "who i am" in statement:
            speak("If you talk then definitely you are a human.")

        elif "assalam walekum" in statement:  # pronunce as as-salamu-alaykum
            speak("Waalaikum asalaam")

        elif "calculate" in statement:
            question = input("Question: ")

            app_id = "RPEQQ8-2EE3Y26WGY"

            client = wolframalpha.Client(app_id)

            res = client.query(question)

            answer = next(res.results).text

            print(answer)
        elif "ask" in statement:
            speak(
                "I can answer to computational and geographical questions and what question do you want to ask now"
            )
            question = takeCommand()
            app_id = "R2K75H-7ELALHR35X"
            client = wolframalpha.Client("R2K75H-7ELALHR35X")
            res = client.query(question)
            answer = next(res.results).text
            speak(answer)
            print(answer)

        elif "search" in statement:
            statement = statement.replace("search", "")
            webbrowser.open_new_tab(statement)
            time.sleep(5)
        elif "where is" in statement:
            statement = statement.replace("where is", "")
            location = statement
            speak("User asked to Locate")
            speak(location)
            webbrowser.open("https://www.google.com / maps / place/" + location + "")

        elif "search" in statement or "play" in statement:
            statement = statement.replace("search", "")
            statement = statement.replace("play", "")
            webbrowser.open(statement)

        elif "start music" in statement or "can i get a song" in statement:
            speak("Here you go with music")

            music_dir = "C:\\295Music"
            songs = os.listdir(music_dir)
            print(songs)
            random = os.startfile(os.path.join(music_dir, songs[1]))

        elif "write a note" in statement:
            speak("What should i write, sir")
            note = takeCommand()
            file = open("jarvis.txt", "a")

            file.write(" :- ")
            file.write(note)
            speak("done")

        elif "open note" in statement:
            speak("Showing Notes")
            file = open("jarvis.txt", "r")
            print(file.read())
            file.seek(0)
            speak(file.read())

        # elif "email to gaurav" in statement:
        #     try:
        #         speak("What should I say?")
        #         content = takeCommand()
        #         to = "freestylewinds777@gmail.com"
        #         sendEmail(to, content)
        #         speak("Email has been sent !")
        #     except Exception as e:
        #         print(e)
        #         speak("I am not able to send this email")

        # elif "send a mail" in statement:
        #     try:
        #         speak("What should I say?")
        #         content = takeCommand()
        #         speak("whome should i send")
        #         to = input()
        #         sendEmail(to, content)
        #         speak("Email has been sent !")
        #     except Exception as e:
        #         print(e)
        #         speak("I am not able to send this email")

        # elif 'send a ' in statement:

        #     try:

        #         speak('What should I say?')

        #         content = takeCommand()

        #         speak('whom should i send')

        #         to = input()

        #         sendEmail(to, content)

        #         speak('Email has been sent !')

        #     except Exception as e:

        #         print(e)

        #         speak('I am not able to send this email')

        elif "weather" in statement:
            api_key = "8ef61edcf1c576d65d836254e11ea420"
            base_url = "https://api.openweathermap.org/data/2.5/weather?"
            speak("whats the city name")
            city_name = takeCommand()
            complete_url = base_url + "appid=" + api_key + "&q=" + city_name
            response = requests.get(complete_url)
            x = response.json()
            if x["cod"] != "404":
                y = x["main"]
                current_temperature = y["temp"]
                current_humidiy = y["humidity"]
                z = x["weather"]
                weather_description = z[0]["description"]
                speak(
                    " Temperature in kelvin unit is "
                    + str(current_temperature)
                    + "\n humidity in percentage is "
                    + str(current_humidiy)
                    + "\n description  "
                    + str(weather_description)
                )
                print(
                    " Temperature in kelvin unit = "
                    + str(current_temperature)
                    + "\n humidity (in percentage) = "
                    + str(current_humidiy)
                    + "\n description = "
                    + str(weather_description)
                )

            else:
                speak(" City Not Found ")
        elif "good bye" in statement or "ok bye" in statement or "stop" in statement:
            speak("your personal assistant G-one is shutting down,Good bye")
            print("your personal assistant G-one is shutting down,Good bye")
            break

        elif "log off" in statement or "sign out" in statement:
            speak(
                "Ok , your pc will log off in 10 sec make sure you exit from all applications"
            )
            subprocess.call(["shutdown", "/l"])
        elif "lock window" in statement:
            speak("locking the device")
            ctypes.windll.user32.LockWorkStation()

        # elif "empty recycle bin" in statement:
        #     winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
        #     speak("Recycle Bin Recycled")

        elif "don't listen" in statement or "stop listening" in statement:
            speak("for how much time you want to stop jarvis from listening commands")
            a = int(takeCommand())
            time.sleep(a)
            print(a)

        elif "restart" in statement:
            subprocess.call(["shutdown", "/r"])


time.sleep(3)
