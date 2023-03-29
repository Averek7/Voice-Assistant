from win32com.client import Dispatch
import winsound
import speech_recognition as sr
import datetime
from datetime import date
import time
import winshell
import pyjokes
import subprocess
import wikipedia
import webbrowser
import os
from googletrans import Translator 
import smtplib
import random
import requests
import wolframalpha
import speedtest
import threading

speak=Dispatch("SAPI.SpVoice")
  
def wishme():
    speak.Speak("Welcome back,Boss")

    hour=int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak.Speak('Good morning')
    elif hour>=12 and hour<18:
        speak.Speak('Good afternoon')
    else:
        speak.Speak('Good evening')

    strtime=datetime.datetime.now().strftime("%H:%M:%S")
    speak.Speak(f"it's {strtime}")
    
    api_key="e3a3b91257d22de4f182fe8e6d1068b3"
    url="https://api.openweathermap.org/data/2.5/weather?"
    compurl=url+"appid="+api_key+"&q="+"Ambikapur"
    res=requests.get(compurl)
    x=res.json()
    if x["cod"]!="401":
        y=x["main"]
        curtemp=y["temp"]
        z=x["weather"]
        weather_ds=z[0]["description"]
        speak.Speak("Temperature is"+str(int(curtemp-273.15))+"degree celcius"
        +"\n description"+str(weather_ds))
    else:
        print("city not found")

    speak.Speak("Ready for your command,Sir")

def takeCommand():
    r=sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening....")
        r.pause_threshold=1
        audio = r.listen(source)
    try:
        print("Recognizing...")
        query = r.recognize_google(audio)
        print(f"User said:{query}\n")
    except:
        print("Please say that again..")
        return "None"
    return query  

def sendEmail(to,content):
    server=smtplib.SMTP("smtp.gmail.com",587)
    server.ehlo()
    server.starttls()
    server.login('singhavi2122@gmail.com','averek710@21')
    server.sendmail('singhavi2122@gmail.com',to,content)
    server.close()

def alarm():
    print(datetime.datetime.now().strftime("%H:%M %p"))
    speak.Speak("At what time,Boss..")
    time=takeCommand()
    times=str(time).upper()
    delst=times.split('.')
    destr="".join(delst)
    print("alarm set for:-",destr)
    speak.Speak("your alarm is on")
    while True:
        if datetime.datetime.now().strftime("%H:%M %p")==destr:
            freq=800
            dur=200
            winsound.Beep(freq,dur)
            break

def check():
    st=speedtest.Speedtest()
    down=str(st.download())
    speak.Speak("Download:-"+ str(int(down))*(10**6)+"mbps")
    up=str(st.upload())
    speak.Speak("Upload:-"+ str(int(up))*(10**6)+"mbps")
    servernames=[]
    st.get_servers(servernames)
    speak.Speak("Ping:-"+ str(st.results.ping))

def newsint():
    main_url = "https://newsapi.org/v1/articles?source=bbc-news&sortBy=top&apiKey=87e58df78d7a432d9bd1621b200b95a6"
    open_bbc_page = requests.get(main_url).json() 
    article = open_bbc_page["articles"] 
    results = []      
    for ar in article: 
        results.append(ar["title"]) 
    for i in range(len(results)): 
        print(i + 1, results[i]) 
    from win32com.client import Dispatch 
    speak = Dispatch("SAPI.Spvoice") 
    speak.Speak(results)

def newsind():
    url = ('http://newsapi.org/v2/top-headlines?'
       'country=in&'
       'apiKey=87e58df78d7a432d9bd1621b200b95a6')
    indnews=requests.get(url).json()
    article=indnews["articles"]
    results=[]
    for ar in article: 
        results.append(ar["title"]) 
    for i in range(len(results)): 
        print(i + 1, results[i]) 
    speak = Dispatch("SAPI.Spvoice") 
    speak.Speak(results)

def ques():
    client=wolframalpha.Client("EGXYWV-WVUTEK56V7")
    res=client.query(query)
    print(next(res.results).text)
    speak.Speak(next(res.results).text)
     
def day():
    day = datetime.datetime.today().weekday() + 1
    Day_dict = {1:'Monday',2:'Tuesday',3:'Wednesday',4:'Thursday',5:'Friday',6:'Saturday',7:'Sunday'} 
    if day in Day_dict.keys(): 
        day_of_the_week = Day_dict[day] 
        print(day_of_the_week) 
        speak.Speak("The day is " + day_of_the_week)

def wakeword(text):
    wake_words=["hey jarvis","ok jarvis","wake up jarvis","are you there jarvis","jarvis"]
    text=text.lower()
    for phrase in wake_words:
        if phrase in text:
            return True
    return False
#------------------------------------------------------------------------------------------------------------------------------
       
if __name__ == "__main__":
    speak.Speak("System initialisation")
    speak.Speak("Loading disks")
    speak.Speak("Security checks")
    speak.Speak("Launching interface")

    wishme()

    while True:
        query=takeCommand().lower()
        if 'wikipedia' in query:
            speak.Speak('Searching wikipedia.. Please wait')
            query=query.replace("wikipedia","")
            result=wikipedia.summary(query,sentences=3)
            speak.Speak(result)
                
        elif 'open google' in query:
            speak.Speak('Okay Boss,Opening.. Please wait')
            webbrowser.open("https://www.google.com")
               
        elif 'open youtube' in query:
            speak.Speak('Okay Boss,Opening.. Please wait')
            webbrowser.open("https://www.youtube.com")
            
        elif 'open instagram' in query:
            speak.Speak('Okay Boss,Opening.. Please wait')
            webbrowser.open("https://www.instagram.com")
        
        elif 'open facebook' in query:
            speak.Speak('Okay Boss,Opening.. Please wait')
            webbrowser.open("https://www.facebook.com")
                
        elif 'open stackoverflow' in query:
            speak.Speak('Okay Boss,Opening.. Please wait')
            webbrowser.open("https://www.stackoverflow.com")
        
        elif "open chrome" in query or "open browser" in query:
            path="C:\\Program Files\\Google\\Chrome\\Application"
                
        elif 'music' in query or 'song' in query:
            speak.Speak("Here you go with music") 
            music_dir="C:\\Users\\Asus\\Music"
            song=os.listdir(music_dir)
            leng=len(song)
            i=random.randint(0,(leng-1))
            os.startfile(os.path.join(music_dir,song[i]))
            
        elif 'play a movie' in query:
            speak.Speak("Here you go with movie") 
            movie_dir="C:\\Users\\ASUS\\Videos"
            play=os.listdir(movie_dir)
            leng=len(play)
            i=random.randint(0,(leng-1))
            os.startfile(os.path.join(movie_dir,play[i]))
                
        elif 'time' in query:
            strtime=datetime.datetime.now().strftime("%H:%M:%S")
            speak.Speak(f'Sir,the time is {strtime}')
            
        elif 'day' in query:
            day()

        elif 'how are you' in query: 
            speak.Speak("I am fine, Thank you") 
            speak.Speak("How are you,Sir") 
  
        elif 'fine' in query : 
            speak.Speak("It's good to know that your fine") 

        elif 'send an email' in query:
            try:
                speak.Speak('To whom?')
                to=takeCommand()
                if "me" in to:
                    speak.Speak('What should i say?')
                    content=takeCommand()
                    sendEmail('singhavanish2003@gmail.com',content)
                    speak.Speak("Boss,your email has been sent!")
                elif "Abhishek" in to:
                    speak.Speak('What should i say?')
                    content=takeCommand()
                    sendEmail('abhig0817@gmail.com',content)
                    speak.Speak("Boss,your email has been sent!")
                elif "Dad" in to:
                    speak.Speak('What should i say?')
                    content=takeCommand()
                    sendEmail('singhdhrmendra@gmail.com',content)
                    speak.Speak("Boss,your email has been sent!")
                elif "bro" in to:
                    speak.Speak('What should i say?')
                    content=takeCommand()
                    sendEmail('singhaditya1982006@gmail.com',content)
                    speak.Speak("Boss,your email has been sent!")
                else:
                    print("no user found")
            except Exception as e:
                print(e)
                speak.Speak("Sorry! Boss email has not been sent ")
                    
        elif 'open python' in query:
            path ="C:\\Users\\ASUS\\AppData\\Local\\Programs\\Python\\Python37\\Lib\\idlelib\\idle.pyw"
            speak.Speak('Okay Boss,Opening.. Please wait')
            os.startfile(path)
                
        elif 'alarm' in query:
            t1 = threading.Thread(target=alarm, name="t1")
            t1.start()

        elif 'open facebook' in query:
            speak.Speak('Okay Boss,Opening.. Please wait')
            webbrowser.open("https://www.facebook.com")

        elif 'current weather' in query:
            api_key="e3a3b91257d22de4f182fe8e6d1068b3"
            url="https://api.openweathermap.org/data/2.5/weather?"
            speak.Speak("Please tell me the city's name")
            city_name=takeCommand()
            compurl=url+"appid="+api_key+"&q="+city_name
            res=requests.get(compurl)
            x=res.json()
            if x["cod"]!="401":
                y=x["main"]
                curtemp=y["temp"]
                curpre=y["pressure"]
                curhum=y["humidity"]
                z=x["weather"]
                weather_ds=z[0]["description"]
                speak.Speak("Temperature is"+str(int(curtemp-273.15))+"degree celcius"
                +"\n atmospheric pressure is "+str(curpre)+"hPa"
                +"\n humidity is"+str(curhum)+"percent"
                +"\n description"+str(weather_ds))
            else:
                speak.Speak("city not found")
                    
        elif "internet speed" in query:
            check()

        elif "international news" in query:
            newsint()

        elif "news" in query:
            newsind()

        elif "shutdown" in query:
            hour=int(datetime.datetime.now().hour)
            if (hour>=21):
                speak.Speak("Good Night!, Sir")
                
            speak.Speak("I'm Off")
            break

        elif "who is" in query or "what is" in query:
            ques()

        elif "good morning" in query or "good afternoon" in query or "good evening" in query or "good night" in query:
            speak.Speak(query)

        elif "thank you" in query or "thanks" in query:
            speak.Speak("Your welcome,Sir")
            
        elif 'who are you' in query or "what's your name" in query or 'what is your name' in query:
            speak.Speak("This is jarvis created by my lazy Boss") 
            
        elif "restart" in query: 
            subprocess.call(["shutdown", "/r"]) 
              
        elif "hibernate" in query: 
            speak.Speak("Hibernating") 
            subprocess.call("shutdown / h")

        elif 'exit system' in query: 
            speak.Speak("Hold On a Sec ! Your system is on its way to shut down") 
            subprocess.call('shutdown / p /f') 
            
        elif 'empty recycle bin' in query: 
            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True) 
            speak.Speak("Recycle Bin Recycled") 
            
        elif 'jokes' in query: 
            speak.Speak(pyjokes.get_joke()) 
              
        elif "calculate" in query:  
            client = wolframalpha.Client("EGXYWV-WVUTEK56V7") 
            indx = query.lower().split().index('calculate')
            query = query.split()[indx + 1:]  
            res = client.query(' '.join(query))  
            answer = next(res.results).text 
            print("The answer is " + answer)  
            speak.Speak("The answer is " + answer)
            
        elif "where is" in query: 
            query = query.replace("where is", "") 
            location = query 
            speak.Speak("User asked to Locate") 
            speak.Speak(location) 
            webbrowser.open("https://www.google.nl/maps/place/"+ location +"")

        elif "search for " in query:
            query=query.replace("search for","")
            search=query
            webbrowser.open("https://google.com/?#q="+search)
    
        elif "play" in query:
            query=query.replace("play","")
            play=query
            webbrowser.open_new("https://youtube.com/results?search_query="+play)        

        elif "sleep" in query or "stop listening" in query:
            speak.Speak("Preparing..")
            while True:
                query=takeCommand().lower()
                if (wakeword(query)==True):
                    speak.Speak("Ready for your command,Boss")
                    break
                else:
                    continue

        elif (wakeword(query)==True):
                speak.Speak("Yes,Sir")
            
        elif "weather forecast" in query:
            client = wolframalpha.Client("EGXYWV-WVUTEK56V7") 
            res=client.query(query)
            print("temperature & description:-"+str(next(res.results).text))
            speak.Speak("temperature & description:-"+str(next(res.results).text))
                
        elif "features" in query:
            speak.Speak('''I can play music, movies, send emails, connected to various websites, do basic mathematics,
            tell current time and date, present weather and whole forcast of the day, locate places over google maps,
            set alarms and reminders, tells national and international news''')
        
        elif "open documents" in query or "open document" in query:
            path=("C:\\Users\\ASUS\\Documents")
            speak.Speak("Launching,Sir")
            os.startfile(path)

        elif "open downloads" in query or "open download" in query:
            path=("C:\\Users\\ASUS\\Downloads")
            speak.Speak("Launching,Sir")
            os.startfile(path)

        elif "open code" in query:
            path=("C:\\Users\\lenovo\\Averek\\Averek.py")
            speak.Speak("Launching,Sir")
            os.startfile(path)

        elif "translate" in query:
            speak.Speak("Please,tell me the text to be translated")
            text=takeCommand().lower()
            translator=Translator() 
            from_lang='en'
            to_lang='hi'
            try: 
                print("Phrase to be Translated :"+ text) 
                text_to_translate=translator.translate(text, src= from_lang, dest= to_lang) 
                text=text_to_translate.text 
                print(text) 
                speak.Speak(str(text))
            except Exception as e: 
                print("Unable to Understand the Input")
                speak.Speak("Unable to Understand the Input")
        
        
        elif "write a note" in query:
            speak.Speak("What should i write,sir")
            note=takeCommand().lower()
            file=open("C:\\Averek\\Jarvis.txt","w")
            speak.Speak("Sir,should i include date and time")
            snfm=takeCommand().lower()
            if "yes" in snfm or "sure" in snfm or "ofcourse" in snfm:
                strtime=datetime.datetime.now().strftime("%H:%M:%S")
                file.write(strtime)
                file.write(":-")
                file.write(note)
                file.write(".")
                speak.Speak("Done")
            else:
                file.write(note)
                speak.Speak("Done")

        elif "show notes" in query:
            speak.Speak("Here are your notes,Boss")
            file=open("C:\\Averek\\Jarvis.txt","r")        
            print(file.read())

        elif "reminder" in query:
            speak.Speak("What shall i remind you about?")
            toremind=takeCommand().lower()
            speak.Speak("And when shall i remind?")
            day=takeCommand().capitalize()
            file=open("C:\\Averek\\Friday.txt","w")
            strtime=datetime.datetime.now().strftime("%H:%M:%S")
            file.write(strtime)
            file.write(",")
            file.write(day)
            file.write(":-")
            file.write(toremind)
            file.write(".")
            speak.Speak("Done")
        
        # elif "updates" in query or "anything new today" in query:
        #     speak.Speak("Let me have a look")
        #     day = datetime.datetime.today().weekday() + 1
        #     Day_dict = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'] 
        #     if day in Day_dict: 
        #         day_of_the_week = Day_dict[day]  
        #     else:
        #         print("Record not found")
                
        #     file=open("C:\\Averek\\Friday.txt","r")
        #     strl=file.readline()
        #     if day_of_the_week in strl and date:
        #         print(strl)
        #         speak.Speak("Should i remove this reminder,Sir")
        #         ded=takeCommand().lower()
        #         if "yes" in ded or "ok" in ded:
        #             file1=open("C:\\Averek\\Friday.txt","w")
        #             for line in file.readline():
        #                 if line.strip("\n")!=file1.readline():
        #                     file1.write(line)
        #                     print(file1.readline())
        #                 else:
        #                     print("data not deleted")
        #         else:
        #             file.close()
        #     else:
        #         print("Invalid record")
            
        
        




        
        
        


