import speech_recognition as sr
import os
import win32com.client
import webbrowser
import openai
import datetime
import requests

openai.api_key = '--Enter your OpenAI API Key--'
weatherapi_key = '--Enter your Weather API Key--'
speaker = win32com.client.Dispatch("SAPI.SpVoice")

chatStr = ""


def chat(query):
    global chatStr
    chatStr += f"User: {query}\nJarvis: "
    conversation = []
    conversation.append({'role': 'system', 'content': chatStr})
    response = openai.ChatCompletion.create(
        model='gpt-3.5-turbo',
        messages=conversation
    )
    chatStr += f"{response.choices[-1].message.content}\n"
    print(chatStr)
    speaker.Speak(response.choices[-1].message.content)
    return (response.choices[-1].message.content)


def ai(prompt):
    conversation = []
    text = f"OpenAI response for Prompt: {prompt} \n *************\n\n"
    conversation.append({'role': 'system', 'content': f"{prompt}"})
    response = openai.ChatCompletion.create(
        model='gpt-3.5-turbo',
        messages=conversation
    )
    # todo: Wrap this inside of a try catch block
    print(response.choices[-1].message.content)
    text += (response.choices[-1].message.content)
    if not os.path.exists("OpenAi"):
        os.mkdir("OpenAi")
    with open(f"OpenAi/{''.join(prompt.split('AI')[1:]).strip()}.txt", "w") as f:
        f.write(text)


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source) 
        try:
            print("Recognising...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Some error occured, Sorry from Jarvis"


# beginning of AI
if __name__ == '__main__':
    print("Hello I am Athena")
    speaker.Speak("Hello I am Athena Desktop Assistant, How can I help you")
    while True:
        print("Listening....")
        query = takeCommand()
        # todo: add more sites
        sites = [["youtube", "https://www.youtube.com/"], ["wikipedia", "https://www.wikipedia.org/"],
                 ["google", "https://www.google.com/"], ["facebook", "https://www.facebook.com/"],
                 ["instagram", "https://www.instagram.com/"]]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                speaker.Speak(f"Opening {site[0]} Sir...")
                webbrowser.open(site[1])
                exit()

        # todo: Add a feature to play a specific song
        if "play music" in query:
            print("Playing music now...")
            speaker.Speak("Playing music now...")
            musicpath = "downfall.mp3"
            os.startfile(musicpath)

        # todo: Give weather stats
        elif " the weather" in query:
            print("Which city weather you want to fetch? ")
            speaker.Speak("Which city weather you want to fetch?")
            print("Listening....")
            user_input = takeCommand()
            weather_data = requests.get(
                f"http://api.weatherapi.com/v1/current.json?key={weatherapi_key}&q={user_input}&aqi=no")
            weather = weather_data.json()['current']['condition']['text']
            temp = weather_data.json()['current']['temp_c']
            humidity = weather_data.json()['current']['temp_c']
            last_updated = weather_data.json()['current']['last_updated']
            print(f"The weather in {user_input} is: {weather}")
            speaker.Speak(f"The weather in {user_input} is: {weather}")
            print(f"The temperature in {user_input} is: {temp}ºC")
            speaker.Speak(f"The temperature in {user_input} is: {temp}º Celsius")
            print(f"The humidity in {user_input} is: {humidity}%")
            speaker.Speak(f"The humidity in {user_input} is: {humidity}%")
            print(f"This report is as of : {last_updated}")
            speaker.Speak(f"This report is as of {last_updated}")

        # todo: Add feature to give the current time
        elif "the time" in query:
            strfTime = datetime.datetime.now().strftime("%H:%M:%S")
            print(f"Sir the time is {strfTime}")
            speaker.Speak(f"Sir the time is {strfTime}")

        elif "postman".lower() in query:
            print("Opening Postman from your desktop..")
            speaker.Speak("Opening Postman from your desktop..")
            os.system(f" /Users/vivek/Desktop/Postman.lnk")

        elif "Using AI".lower() in query.lower():
            speaker.Speak("Generating using AI. Please wait")
            ai(prompt=query)
            speaker.Speak("Done Sir, Your file is saved in OpenAi Directory. Is there anything else you would like me "
                          "to do for you?")
        elif "stop".lower() in query.lower():
            exit()
        elif "reset chat".lower() in query.lower():
            chatStr = ""
        else:
            chat(query)

        # speaker.Speak(query)
