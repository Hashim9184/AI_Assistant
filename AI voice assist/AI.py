import speech_recognition as sr
import pyttsx3
import pywhatkit
import datetime
import wikipedia
import requests
import os
import wolframalpha #calculate computational and geographical data
import webbrowser as wb
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
import screen_brightness_control as sbc
from tkinter import *    # GUI modules from here

window =Tk()
window.title("Max")
canvas = Canvas(window,height=500,width=500, bg = 'black')
canvas.pack(expand = YES, fill = BOTH, anchor=CENTER,)
my_image = PhotoImage(file='./MAX/jarvi.gif')
canvas.create_image(250, 200, image = my_image, anchor = CENTER)

listener = sr.Recognizer()
engine = pyttsx3.init()
voices= engine.getProperty('voices') #getting details of current voice
engine.setProperty('voice', voices[0].id, )


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        talk("Good Morning Sir!")
    elif hour>=12 and hour<18:
        talk("Good Afternoon")

    else:
        talk("Good Evening")

    talk("I am Max. Please tell me how can I help you")

def talk(audio):
    engine.say(audio)
    engine.runAndWait()
def take_command():
    wishMe()
    try:
        with sr.Microphone() as data_taker:
            print("Say Somethig")
            voice = listener.listen(data_taker)
            instruct = listener.recognize_google(voice)
            instruct = instruct.lower()
        if 'Max' in instruct:
            instruct = instruct.replace('Max', '')
            print(instruct)
    except:
        pass
    return instruct


def run_Max():
    instruct = take_command()
    if 'play' in instruct:
        song = instruct.replace('play', '')
        talk('playing' + song)
        pywhatkit.playonyt(song)

    elif 'time' in instruct:
        time = datetime.datetime.now().strftime('%I: %M')
        print(time)
        talk('current time is' + time)
    
    elif 'calculate' in instruct:
        talk('I can answer to computational and geographical questions  and what question do you want to ask now')
        question=instruct.replace('Hey','')
        print(question)
        app_id="AKHT6H-4YAYH4GYQ7"
        client = wolframalpha.Client('R2K75H-7ELALHR35X')
        res = client.query(question)
        answer = next(res.results).text
        talk(answer)
        print(answer)
    
    elif "weather" in instruct:
        api_key="8ef61edcf1c576d65d836254e11ea420"
        base_url="https://api.openweathermap.org/data/2.5/weather?"
        city_name= instruct.replace("weather of", "")
        print(city_name)
        complete_url=base_url+"appid="+api_key+"&q="+city_name
        response = requests.get(complete_url)
        x=response.json()
        if x["cod"]!="404":
            y=x["main"]
            current_temperature = y["temp"]
            current_humidiy = y["humidity"]
            z = x["weather"]
            weather_description = z[0]["description"]
            talk(" Temperature in kelvin unit is " +
                    str(current_temperature) +
                    "\n humidity in percentage is " +
                    str(current_humidiy) +
                    "\n description  " +
                    str(weather_description))
            print(" Temperature in kelvin unit = " +
                    str(current_temperature) +
                    "\n humidity (in percentage) = " +
                    str(current_humidiy) +
                    "\n description = " +
                    str(weather_description))

    elif  '''today's headline''' in instruct:
        news = wb.open_new_tab("https://timesofindia.indiatimes.com/home/headlines")
        talk('Here are some headlines from the Times of India,Happy reading')
        time.sleep(6)

    elif 'set brightness' and '50' in instruct:
            sbc.fade_brightness(50)
            print(sbc.get_brightness())
            print("Brightness set to 50 successfully")

    elif 'set brightness' and '100' in instruct:
        sbc.fade_brightness(100)
        print(sbc.get_brightness())
        print("I've made the screen brighter")

    elif 'set brightness' and '25' in instruct:
            sbc.set_brightness(25)
            print("Brightness set to 25 successfully")

    elif 'tell me something about' in instruct:
        thing = instruct.replace('tell me something about', '')
        print(thing)
        info = wikipedia.summary(thing, 2)
        print(info)
        talk(info)
    
    elif 'set volume to half' in instruct:
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        # Get current volume 
        currentVolumeDb = volume.GetMasterVolumeLevel()
        print(currentVolumeDb)
        volume.SetMasterVolumeLevel(currentVolumeDb - 6.0, None)
        print(currentVolumeDb)
        talk("Volume set to half")

    elif 'mute' in instruct:
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        # Get current volume 
        currentVolumeDb = volume.GetMasterVolumeLevel()
        volume.SetMasterVolumeLevel(currentVolumeDb - 29.0, None)
        print("Device mutted sucessfully")

    elif "where is" in instruct:
        query = instruct.replace("where is", "")
        location = instruct
        talk("User asked to Locate")
        talk(query)
        wb.open("https://www.google.com/maps/place/"+ query +"")
    elif 'brave' in instruct:
        talk("Opening")
        talk("Brave Browser")
        os.startfile("brave")
    elif 'word' in instruct:
        talk("Opening")
        talk("MS Word")
        os.startfile("C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE")
    elif 'gmail' in instruct:
        talk("Opening")
        talk("Gmail")
        wb.open("https://mail.google.com/mail/u/0/#inbox")

    elif 'command prompt' in instruct:
        talk("Opening")
        talk("Command Prompt")
        os.startfile("cmd")
    elif ('.com') in instruct:
        talk(instruct)
        Chrome = ("C://Program Files/Google/Chrome/Application/chrome.exe %s")
        wb.get(Chrome).open('http://www.'+instruct)

    elif 'who are you' in instruct:
        talk('I am your personal Assistant Max')
        print('I am your personal Assistant Max')
    elif 'what can you do for me' in instruct:
        talk('I can play songs, tell time, and help you go with wikipedia')
        print('I can play songs, tell time, and help you go with wikipedia')
    elif "who made you" in instruct or "who created you" in instruct or "who discovered you" in instruct:
        talk("I was built  by a team of developers")
        print("I was built  by a team of developers")
    else:
        talk('I did not understand, can you repeat again')

btn1 =Button(window,text = "Initialize Max", command= run_Max)
btn1.pack()
window.mainloop()
