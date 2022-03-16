# from typing_extensions import Required
import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import os
import subprocess
# from googlesearch import search
import smtplib
import shutil
import requests
import webbrowser as wb
import time
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import screen_brightness_control as sbc
# import face_recognition as fr
# import cv2
# import numpy as np
# from wifi import createNewConnection
import pyaudio
from wikipedia import search


engine = pyttsx3.init('sapi5')
voices= engine.getProperty('voices') #getting details of current voice
# print(voices)
engine.setProperty('voice', voices[0].id)
def speak(audio):
    engine.say(audio)
    engine.runAndWait() 

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning Sir!")
    elif hour>=12 and hour<18:
        speak("Good Afternoon")

    else:
        speak("Good Evening")

    speak("I am Jarvis. Please tell me how can I help you")
 
def takeCommand():
     r= sr.Recognizer()
     with sr.Microphone() as source:
            print("Listning...")
            r.pause_threshold = 1
            audio = r.listen(source)
    
     try:
        print("Recognizing...")    
        query = r.recognize_google(audio, language='en-in') #Using google for voice recognition.
        print(f"User said: {query}\n")  #User query will be printed.

     except Exception as e:
        print("Say that again please...")   #Say that again will be printed in case of improper voice 
        return "None" #None string will be returned
     return query

def sendEmail(to, content):
                server = smtplib.SMTP('smtp.gmail.com', 587)
                server.ehlo()
                server.starttls()

                server.login('hashim.shaikh09184@gmail.com', 'siraj9184')
                server.sendmail('hashim.shaikh09184@gmail.com', to, content)
                server.close()

# def createNewConnection(name, SSID, password):
#     config = """<?xml version=\"1.0\"?>
# <WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">
#     <name>"""+name+"""</name>
#     <SSIDConfig>
#         <SSID>
#             <name>"""+SSID+"""</name>
#         </SSID>
#     </SSIDConfig>
#     <connectionType>ESS</connectionType>
#     <connectionMode>auto</connectionMode>
#     <MSM>
#         <security>
#             <authEncryption>
#                 <authentication>WPA2PSK</authentication>
#                 <encryption>AES</encryption>
#                 <useOneX>false</useOneX>
#             </authEncryption>
#             <sharedKey>
#                 <keyType>passPhrase</keyType>
#                 <protected>false</protected>
#                 <keyMaterial>"""+password+"""</keyMaterial>
#             </sharedKey>
#         </security>
#     </MSM>
# </WLANProfile>"""
#     command = "netsh wlan add profile filename=\""+name+".xml\""+" interface=Wi-Fi"
#     with open(name+".xml", 'w') as file:
#         file.write(config)
#     os.system(command)
 
# # function to connect to a network   
# def connect(name, SSID):
#     command = "netsh wlan connect name=\""+name+"\" ssid=\""+SSID+"\" interface=Wi-Fi"
#     os.system(command)
 
# # function to display avavilabe Wifi networks   
# def displayAvailableNetworks():
#     command = "netsh wlan show networks interface=Wi-Fi"
#     os.system(command)
 
 
# # display available netwroks
# displayAvailableNetworks()
 
# # input wifi name and password
# name = input("Name of Wi-Fi: ")
# password = input("Password: ")
 
# # establish new connection
# createNewConnection(name, name, password)
 
# # connect to the wifi network
# connect(name, name)
# print("If you aren't connected to this network, try connecting with the correct password!")

if __name__ == "__main__":
    wishMe()
    # while True:
    if 1:
        query = takeCommand().lower() #Converting user query into lower case

        # Logic for executing tasks based on query
        if 'wikipedia' in query:  #if wikipedia found in the query then this block will be executed
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2) 
            speak("According to Wikipedia")
            print(results)
            speak(results)

            #Opening browser
        elif 'open' in query and 'browser':
            queryArray = query.split()
            queryToSearch = queryArray[queryArray.index('open')+1]
            searchResult= search(queryToSearch, tld="com", num=1, start=0, stop=1, pause=2).__next__()
            os.startfile(searchResult)

             # Shutting down system 
        elif 'shutdown' in query:
            print("Shutting down the computer")
            speak("Shutting the computer")
            os.system('shutdown /s /t 30')

            # Opening application from sys
        elif 'msexcel' in query:
            pyttsx3.speak("Opening")
            pyttsx3.speak("MICROSOFT EXCEL")
            print(".")
            print(".")
            os.startfile("Excel")

        elif 'word' in query:
            pyttsx3.speak("Opening")
            pyttsx3.speak("MICROSOFT WORD")
            os.startfile("WINWORD")

        elif 'brave' in query:
            speak("Opening")
            speak("Brave Browser")
            os.startfile("brave")

        elif 'command prompt' in query:
            speak("Opening")
            speak("Command Prompt")
            os.startfile("cmd")

        elif 'time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")    
            print(strTime)
            speak(f"Sir, the time is {strTime}") 
        
        elif 'email' in query:
                try:
                    speak("Sending Email")
                    content = takeCommand()
                    to = 'noaman@gmail.com'
                    sendEmail(to, content)
                    speak("Email has been sent!")
                except Exception as e:
                    print(e)
                    speak("Sorry my friend. I am not able to send this email")


        elif 'delete a folder' in query:
            try:
                speak("deleting folder")
                os.rmdir("C:\\New folder")
                speak("deleted folder successfully")
            except Exception as d:
                print(d)
                speak("File not deleted")

        # elif 'lol' in query:
        #     source = 'C://'
        #     destination = 'C://New folder'
        #     shutil.copy(source, destination)
        #     speak("Copied")

        elif "weather" in query:
            api_key="8ef61edcf1c576d65d836254e11ea420"
            base_url="https://api.openweathermap.org/data/2.5/weather?"
            speak("whats the city name")
            city_name=takeCommand()
            complete_url=base_url+"appid="+api_key+"&q="+city_name
            response = requests.get(complete_url)
            x=response.json()
            if x["cod"]!="404":
                y=x["main"]
                current_temperature = y["temp"]
                current_humidiy = y["humidity"]
                z = x["weather"]
                weather_description = z[0]["description"]
                speak(" Temperature in kelvin unit is " +
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

            else:
                speak(" City Not Found ")

        elif 'news' in query:
            news = wb.open_new_tab("https://timesofindia.indiatimes.com/home/headlines")
            speak('Here are some headlines from the Times of India,Happy reading')
            time.sleep(6)

        elif "camera" in query or "take a photo" in query:
            subprocess.run('start microsoft.windows.camera:', shell=True)

        elif 'close camera' in query:
            subprocess.run('Taskkill /IM WindowsCamera.exe /F', shell=True)

        

        elif 'set volume to half' in query:
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(
            IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            # Get current volume 
            currentVolumeDb = volume.GetMasterVolumeLevel()
            print(currentVolumeDb)
            volume.SetMasterVolumeLevel(currentVolumeDb - 6.0, None)
            print(currentVolumeDb)

        elif 'mute' in query:
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(
            IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            # Get current volume 
            currentVolumeDb = volume.GetMasterVolumeLevel()
            volume.SetMasterVolumeLevel(currentVolumeDb - 29.0, None)

        # scan available Wifi networks
            # os.system('cmd /c "netsh wlan show networks"')
 
#         elif 'wifi' or 'connect to wifi' in query:
# # input Wifi name
#             name_of_router = input('Enter Name/SSID of the Wifi Network you wish to connect to: ')
# # connect to the given wifi network
#             os.system(f'''cmd /c "netsh wlan connect name={name_of_router}"''')
 
#             print("If you're not yet connected, try connecting to a previously connected SSID again!")

        # elif 'new wifi':
        #     Required(createNewConnection())


        elif ('install') in query:
            stopwords = ['install']
            querywords = query.split()
            resultwords  = takeCommand()
            result = ' '.join(resultwords)
            speak('installing '+result)
            os.system('python -m pip install ' + result)

        elif 'start music' and 'start' in query:   
            stopwords = ['start']
            querywords = query.split()
            resultwords  = takeCommand()
            result = ' '.join(resultwords)
            os.system('start ' + result)
            speak('starting '+result)

        elif ('stop music') and ('stop') in query:
            stopwords = ['stop']
            querywords = query.split()
            resultwords  = [word for word in querywords if word.lower() not in stopwords]
            result = ' '.join(resultwords)
            os.system('taskkill /im ' + result + '.exe /f')
            speak('stopping '+result)

        elif ('google maps') in query:
            stopwords = ['google', 'maps']
            querywords = query.split()
            resultwords  = takeCommand()
            result = ' '.join(resultwords)
            Chrome = ("C://Program Files/Google/Chrome/Application/chrome.exe %s")
            wb.get(Chrome).open("https://www.google.be/maps/place/"+result+"/")
            speak(result+'on google maps')

        elif ('.com') in query :
            speak('Opening' + query)
            Chrome = ("C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s")
            wb.get(Chrome).open('http://www.'+query)

        elif 'set brightness' and '75' in query:
            current_brightness = sbc.get_brightness()
        
            # get the brightness of the primary display
            primary_brightness = sbc.get_brightness(display=0)

            sbc.set_brightness(50)

            sbc.set_brightness(75, display=0)
 
            print(sbc.get_brightness())

        elif 'set brightness' and '50' in query:
            sbc.fade_brightness(50)
            print(sbc.get_brightness())

        elif 'set brightness' and '100' in query:
            sbc.fade_brightness(100, increment = 10)
            print(sbc.get_brightness())

        elif 'set brightness' and '25' in query:
            sbc.set_brightness(25, display=0)

        elif 'your name' in query:
            speak('My name is Jarvis, at your service sir')

        elif '*' in query:
            speak('Be polite please')

        elif 'how are you' or ('and you') in query or ('are you okay') in query:
            speak('Fine thank you')

        elif 'jarvis' in query:
            speak('Yes Sir?', 'What can I doo for you sir?')

        elif ('thanks') in query or ('tanks') in query or ('thank you') in query:
            speak('You are wellcome', 'no problem')

        elif ('hello') in query or ('hi') in query:
            speak('Wellcome to Jarvis virtual intelligence project. At your service sir.')

        elif ('goodbye') in query:                          
            speak('Goodbye Sir', 'Jarvis powering off in 3, 2, 1, 0')

        elif 'sleep' or 'sleep mode':
            speak('Good night')
            os.system('rundll32.exe powrprof.dll,SetSuspendState 0,1,0')
            
        # elif 'wifi' in query:
        #     createNewConnection()