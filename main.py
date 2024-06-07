# IMPORT LIBRARIES
from win32com.client import Dispatch
import speech_recognition as sr
from datetime import datetime
import os
import webbrowser
import wikipedia
import openai
import requests
import my_api_keys  # import api from another file
import pywhatkit
from smtplib import SMTP

# FUNCTIONS


def speak(text):
    """ This function takes string or sentence as an input. And speaks it."""
    voice = Dispatch('SAPI.spvoice')
    voice.speak(text)


def takecommand():
    """ This function takes command from user(listen user's voice), recognize it using google recognizer. And convert
    it into string or sentence"""
    r=sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language='en-in')
            print(f"User: {query}")
            return query
        except Exception as error:
            print("TimeOut sir. Try again")
            speak("TimeOut sir. Try again")
            quit()


def ai(command):
    openai.api_key = my_api_keys.openai_api

    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=command,
        temperature=1,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    if not os.path.exists("openai files"):
        os.mkdir("openai files")

    file_name=command.split("using ai")[0]
    with open(f"openai files/{file_name.strip()}.txt",'w') as f:
        f.write(response["choices"][0]["text"])


chatstr = ''


def chat(command):
    openai.api_key = my_api_keys.openai_api

    global chatstr
    chatstr += f"User: {command}\nJarvis: "

    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=chatstr,
        temperature=1,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    chatstr += f"{response['choices'][0]['text'].strip()}\n"
    print(f"Jarvis: {response['choices'][0]['text'].strip()}")
    speak(response['choices'][0]['text'])


def news(about_news):
    api = my_api_keys.news_api
    date = datetime.today().strftime('%y-%m-%d')
    speak("Working on it Sir!")
    data = requests.get(f'https://newsapi.org/v2/everything?q={about_news}&from={date}&sortBy=publishedAt&apiKey={api}')
    json_data = data.json()

    if json_data['status']=='ok':
        if json_data['totalResults']==0:
            print('Sorry Sir. Currently I am not able to find news. Try later.')
            speak('Sorry Sir. Currently I am not able to find news. Try later.')
        elif json_data['totalResults']!=0:
            art=json_data['articles']
            speak(f"Here are some top breaking news about {about_news}.")
            for i in range(5):
                print(art[i]['title'])
                speak(art[i]['title'])
    else:
        print('Sorry Sir. There is some problems while fetching news. Try after some time.')
        speak('Sorry Sir. There is some problems while fetching news. Try after some time.')


def weather(city):
    api = my_api_keys.weather_api
    url = f"http://api.openweathermap.org/data/2.5/weather?appid={api}&q={city}"

    response = requests.get(url).json()
    if response['cod']=='404':
        speak("Sorry Sir, City not found.")
        print("Sorry Sir, City not found.")
        return
    else:
        celsius = int(response['main']['temp']-273.15)
        humidity = response['main']['humidity']
        description = response['weather'][0]['description']
        wind_speed = response['wind']['speed']

        print(f"General weather is: {description}.\nTemperature is: {celsius}°C.\nHumidity is: {humidity}%.\nAnd wind speed is: {wind_speed} m/s")
        speak(f"In {city}, General weather is {description}. Temperature is {celsius} Degree celsius. Humidity is {humidity} percentage. And wind speed is {wind_speed} meter per second")


def wp(name):
    try:
        speak("Working on it Sir.")
        res = wikipedia.summary(name, sentences=2)
        speak("According to Wikipedia.")
        print(res)
        speak(res)

    except wikipedia.exceptions.DisambiguationError as e:
        res = wikipedia.summary(e.options[0], sentances=2)
        speak("According to Wikipedia.")
        print(res)
        speak(res)

    except wikipedia.exceptions.PageError as e:
        speak(f"Sorry Sir, I am not able to search about this. But I am redirecting you to google.")
        pywhatkit.search(name)


def play_video(video):
    print("Playing related video on youtube...")
    pywhatkit.playonyt(video)


def msg_on_whats(number, message):
    speak("Working On It Sir..")
    pywhatkit.sendwhatmsg_instantly(number, message)


def search_on_google(topic):
    speak("Working On It Sir..")
    try:
        pywhatkit.search(topic)
    except:
        speak("Sorry sir, I am not able to search this on google!")


def send_email():
    emails = my_api_keys.email_pass
    print([i[0] for i in emails])
    speak(f"Sir, which email address do you want to use to send mail? {emails[0][0]} or {emails[1][0]}")
    use_email = takecommand()
    if 'first'.lower() in use_email.lower() or 'aryanparwani'.lower() in use_email.lower().replace(' ',''):
        sender_email = emails[0][0]
        password = emails[0][1]
    elif 'second'.lower() in use_email.lower() or 'jarvis'.lower() in use_email.lower():
        sender_email = emails[1][0]
        password = emails[1][1]
    else:
        speak("Somthing went wrong!")
        return

    speak("Now, tell me the receiver's email address.")
    receiver_email = takecommand().lower()
    receiver_email = receiver_email.replace('gmail.com','@gmail.com').replace(' ','')
    if '@@' in receiver_email:
        receiver_email = receiver_email.replace('@@', '@')

    speak("What message do you want to send through mail?")
    msg = takecommand()

    speak("O.K. Working on it sir..")

    with SMTP('smtp.gmail.com',587) as smtp:
        smtp.starttls()
        try:
            smtp.login(sender_email, password)
            smtp.sendmail(sender_email, receiver_email, msg)
            speak("Email sent successfully")
        except:
            print("There is some problems while sending email. Check given email details or try later.")
            speak("There is some problems while sending email. Check given email details or try later.")


if __name__ == '__main__':
    while True:
        take = takecommand()

        sites=[['Youtube', 'https://youtube.com'],
               ['Google','https://google.com'],
               ['chat gpt','https://chat.openai.com'],
               ['wikipedia','https://wikipedia.com'],
               ['stack overflow','https://stackoverflow.com']
               ]
        web_opened=False
        for site in sites:
            if f"open {site[0]}".lower() in take.lower():
                speak(f"Opening {site[0]} sir")
                webbrowser.open(site[1])
                web_opened=True

        if web_opened == True:
            continue
        else:
            pass

        if ('current time'.lower() in take.lower()) or ('what time'.lower() in take.lower()):
            hour=datetime.now().strftime("%H")
            min=datetime.now().strftime("%M")
            print(f"its {hour} hours and {min} minutes.")
            speak(f"its {hour} hours and {min} minutes.")

        elif 'new python file'.lower() in take.lower():
            speak("O.K., Tell me file name.")
            new_file = takecommand()
            open(f'{new_file}.py','w')
            speak("File created Successfully Sir.")

        elif 'top news'.lower() in take.lower() or 'breaking news'.lower() in take.lower():
            speak("O.K., Tell me a title, About which you want to hear news.")
            about_news = takecommand()
            news(about_news)

        elif 'Weather'.lower() in take.lower():
            speak("O.K., Tell me the city name, of which you want to know the weather.")
            city = takecommand()
            weather(city)

        elif 'who is'.lower() in take.lower() or 'wikipedia'.lower() in take.lower():
            l = ['jarvis', 'wikipedia' 'search']
            for i in l:
                take = take.lower().replace(i, '').strip()
            wp(take)

        elif 'play'.lower() in take.lower() and 'on youtube'.lower() in take.lower():
            speak("What do you want to play Sir?")
            video = takecommand()
            speak("Working on it Sir.")
            play_video(video)

        elif 'search'.lower() in take.lower() and 'on google'.lower() in take.lower():
            speak("Sir, What do want to search on google?")
            topic = takecommand()
            search_on_google(topic)

        elif 'send'.lower() in take.lower() and 'email'.lower() in take.lower():
            send_email()

        elif 'message'.lower() in take.lower() and 'whatsapp'.lower() in take.lower():
            name_numbers = my_api_keys.names_with_number
            in_name = False
            for i in name_numbers:
                if i.lower() in take.lower():
                    in_name = True
                    speak(f"What message do you want to send, to {i}")
                    message = takecommand()
                    msg_on_whats(name_numbers[i], message)

            wrong_name = False
            if in_name == False:
                speak("Tell me the name to which you want to send message")
                name = takecommand()
                for i in name_numbers:
                    if i.lower() == name.lower():
                        speak(f"What message do you want to send, to {i}")
                        message = takecommand()
                        msg_on_whats(name_numbers[i], message)
                        wrong_name = False
                        break
                    else:
                        wrong_name = True

                if wrong_name == True:
                    print(f"Sorry sir, i am not able to send message to {name}")
                    speak(f"Sorry sir, i am not able to send message to {name}")

        elif 'using ai'.lower() in take.lower():
            speak("Working on it Sir...")
            ai(take.lower().replace("using ai",''))
            print("Done sir. You can check your answer in openai folder after program completes.")
            speak("Done sir. You can check your answer in open a.i. folder after program completes. Here is path:")
            print(f"C://Users//aryan//PycharmProjects//JARVIS_AI//openai files//{take.lower().split('using ai')[0]}")

        elif 'reset chat'.lower() in take.lower():
            chatstr = ''
            speak("Chat reset successfully Sir.")

        elif 'JARVIS quit'.lower() in take.lower() or 'jarvis stop'.lower() in take.lower():
            print("Stopped")
            speak("o.k. sir. Quitting")
            exit()

        elif ('shutdown' in take.lower().replace(' ','') or 'turn off' in take.lower()) and ('pc' in take.lower() or 'computer' in take.lower()):
            speak("Are you sure? Do you want to shutdown your computer?")
            ans = takecommand()
            if 'yes' in ans.lower():
                try:
                    speak("O.K. Your computer is going to be shut down in 5 seconds!")
                    pywhatkit.shutdown(5)
                except:
                    print("Something Went Wrong!")
                    speak("Something Went Wrong!")

            elif 'no' in ans.lower():
                speak("O.K. sir, stopping shut down!")

        else:
            chat(take.lower())