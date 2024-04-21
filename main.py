import speech_recognition as sr
import win32com.client
speaker = win32com.client.Dispatch("SAPI.SpVoice")
import webbrowser
import pyautogui
import os
import time
import datetime
import openai
from config import apikey
from config import newsAPIkey
import requests

#written by priyanshu yadav 2ndET
def googleSearch(query):

    desired_part = query.split('Google', 1)
    webbrowser.open("https://google.com")
    time.sleep(2)
    pyautogui.hotkey('ctrl', 'l')
    pyautogui.write(desired_part[1], interval=0.1)

    pyautogui.press('enter')


def youtubesearch(query):
    desired_part = query.split('YouTube', 1)
    webbrowser.open("https://youtube.com")
    time.sleep(2)
    pyautogui.hotkey('/')
    pyautogui.write(desired_part[1], interval=0.1)

    pyautogui.press('enter')


def searchsongonyoutube(query):
    desired_part = query.split('YouTube', 1)
    webbrowser.open("https://youtube.com")
    time.sleep(2)
    pyautogui.hotkey('/')
    pyautogui.write(desired_part[1], interval=0.1)

    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('enter')




newsAPI = newsAPIkey
def news():
    main_url="https://newsapi.org/v2/top-headlines?country=in&category=business&apiKey="+newsAPI
    news=requests.get(main_url).json()
    article=news['articles']

    news_article=[]
    for arti in article:
        news_article.append(arti['title'])

    for i in range(5):
        speaker.Speak(f"News {i+1}")
        affairs = news_article[i]
        speaker.Speak(affairs)

chatStr = ""

def chat(query):
    global chatStr
    print(chatStr)

    openai.api_key = apikey
    chatStr += f"\n Aastha: {query} \n AI:"

    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=chatStr,
        temperature=1,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )

    speaker.Speak(response["choices"][0]["text"])
    chatStr += f"{response['choices'][0]['text']}"
    return response['choices'][0]['text']

    with open(f"Openai/{''.join(prompt.split('intelligence')[1:]).strip()}.txt", "w") as f:
        f.write(text)



def ai(prompt):
    openai.api_key = apikey
    text=f"OpenAI response for Prompt: {prompt} \n ************************ \n \n"

    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=prompt,
        temperature=1,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )

    # print(response["choices"][0]["text"])
    text += response["choices"][0]["text"]
    if not os.path.exists("Openai"):
        os.mkdir("Openai")

    with open(f"Openai/{''.join(prompt.split('intelligence')[1:]).strip()}.txt", "w") as f:
        f.write(text)



def takeCommand():
    r= sr.Recognizer()
    with sr.Microphone() as source:
        # r.pause_threshold = 1
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Some error occured, I am Sorry"


speaker.Speak("Hello I am your A.I. Friend, Krish")
while True:
    print("Listening...")
    query = takeCommand()
    sites=[["youtube","https://youtube.com"],["wikipedia","https://wikipedia.com"], ["google","https://google.com"],["Amazon","https://Amazon.com"],["gmail","https://mail.google.com/mail/u/0/#inbox"], ["google meet","https://meet.google.com/?pli=1"]]
    for site in sites:
        if f"Open {site[0]}".lower() in query.lower():
            # speaker.Speak(f"Opening {site[0]} Ma'am'....")
            webbrowser.open(site[1])

    if f"Play song on youtube".lower() in query.lower():
        searchsongonyoutube(query)

    elif f"Play song".lower() in query.lower():
        # speaker.speak("tell me the song name")
        # print("Listening song name...")
        # query = takeCommand()
        desired_part = query.split('song', 1)
        speaker.speak(f"playing {desired_part[1]} in spotify")
        os.system("spotify")
        time.sleep(2)
        pyautogui.hotkey('ctrl', 'l')
        pyautogui.write(desired_part[1], interval=0.1)

        for key in ['enter', 'pagedown', 'tab', 'enter', 'enter']:
            time.sleep(2)
            pyautogui.press(key)

    elif f"the time" in query:
        strfTime = datetime.datetime.now().strftime("%H:%M")
        speaker.speak(f"Mam the time is {strfTime}")

    elif f"Open Camera".lower() in query.lower():
        os.system("start microsoft.windows.camera:")

    elif f"Open This pc".lower() in query.lower():
        os.system('explorer shell:MyComputerFolder')

    elif f"Open Volume Documents".lower() in query.lower():
        os.system("explorer shell:MyComputerFolder\\D")

    elif f"Open Volume D".lower() in query.lower():
        os.system("explorer shell:MyComputerFolder\\D:")


    elif f"using artificial intelligence".lower() in query.lower():
        ai(prompt=query)

    elif f"Krish Quit".lower() in query.lower():
        speaker.Speak('Raadhei raadhei Aastha')
        exit()

    elif f"reset chat".lower() in query.lower():
        chatStr=""

    elif f"What is your name".lower() in query.lower():
        speaker.Speak(f"My name is Krish")

    elif f"Current affairs".lower() in query.lower():
        news()

    elif f"Search on google".lower() in query.lower():
        googleSearch(query)

    elif f"Search on youtube".lower() in query.lower():
        youtubesearch(query)



    else:
        print('Chating....')
        chat(query)