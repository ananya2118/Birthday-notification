import os
import pandas as pd
import datetime
from plyer import notification
import pyttsx3
engine=pyttsx3.init('sapi5')
voices=engine.getProperty('voices')
engine.setProperty('voice',voices[0].id)
def speak(audio):
    engine.say(audio)
    engine.runAndWait()
def notification1(title,msg):
        notification.notify(
           title=title,
           message=msg,
           app_icon="C:\\Users\\Ananya N\\Documents\\python-prog\\wishbirthday\\cake2.ico",
           timeout=10 
        )
def notification2(title,msg):
        notification.notify(
           title=title,
           message=msg,
           app_icon="C:\\Users\\Ananya N\\Documents\\python-prog\\wishbirthday\\ann.ico",
           timeout=10 
        )
df=pd.read_excel("C:\\Users\\Ananya N\\Documents\\python-prog\\wishbirthday\\exc.xlsx")
today=datetime.datetime.now().strftime("%d|%m")
for index,item in df.iterrows():
    bd=item["birthday"]
    if(today==bd):
        a=item["name"]
        notification1("Birthday alert","Its "+a+"'s birthday today.")
        speak("Ananya !! Its "+a+"'s birthday today")
for index,item in df.iterrows():
    an=item["anniversary"]
    if(today==an):
        a=item["names"]
        notification2("Anniversary alert","Its "+a+"'s anniversary today.")
        speak("Ananya !! Its "+a+"'s anniversary today")

        


