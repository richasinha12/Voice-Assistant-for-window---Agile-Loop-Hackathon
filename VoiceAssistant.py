import subprocess
import wolframalpha
import pyttsx3
import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import time
import requests
import shutil
from twilio.rest import Client
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText



engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)
engine.setProperty('rate', 150)


def speak(audio):
	engine.say(audio)
	engine.runAndWait()

def wishMe():
	hour = int(datetime.datetime.now().hour)
	if hour>= 0 and hour<12:
		speak("Good Morning Sir !")

	elif hour>= 12 and hour<18:
		speak("Good Afternoon Sir !") 

	else:
		speak("Good Evening Sir !") 

	assname =("Jarvis 1 point o")
	speak("I am your Assistant")
	speak(assname)
	

def username():
	speak("What should i call you sir")
	uname = takeCommand()
	speak("Welcome Mister")
	speak(uname)
	columns = shutil.get_terminal_size().columns
	
	print("#####################".center(columns))
	print("Welcome Mr.", uname.center(columns))
	print("#####################".center(columns))
	
	speak("How can i Help you, Sir")

def takeCommand():
	
	r = sr.Recognizer()
	
	with sr.Microphone() as source:
		
		print("Listening...")
		# r.pause_threshold = 1
		audio = r.listen(source)

	try:
		print("Recognizing...") 
		query = r.recognize_google(audio, language ='en-in')
		print(f"User said: {query}\n")

	except Exception as e:
		print(e) 
		print("Unable to Recognize your voice.") 
		return "None"
	
	return query

def sendEmail(subject, body, to_email):
	server = smtplib.SMTP('smtp.gmail.com', 587)
	server.ehlo()
	server.starttls()
	msg = MIMEMultipart()
	
    # Email sending credentials
	from_email = "swiftplanai@gmail.com"
	app_password = "wici gxpi sjwn znkj"
	
    # Create the container email message.
	msg['From'] = from_email
	msg['To'] = to_email
	msg['Subject'] = subject
	
    # Attach the body with the msg instance
	msg.attach(MIMEText(body, 'plain'))
	

	# Enable low security in gmail
	server.login(from_email, app_password)
	text = msg.as_string()
	server.sendmail(from_email, to_email , text)
	print('Email sent')
	speak('Email Sent at: {0}'.format(to_email))
	server.close()
	

# testing functions
# sendEmail(subject='Hello World', body='My name is Ali Khan. I am here to teach LLM', to_email='Alitariqdev@gmail.com')
# username()


if __name__ == '__main__':
	clear = lambda: os.system('cls')
	
	# This Function will clean any
	# command before execution of this python file
	# clear()
	# wishMe()
	# username()
	
	while True:
		
		query = takeCommand().lower()
		print(query)
		# All the commands said by user will be 
		# stored here in 'query' and will be
		# converted to lower case for easily 
		# recognition of command



		# tested
		if 'wikipedia' in query:
			speak('What you want to search on wikipedia')
			query = takeCommand()
			speak('Searching Wikipedia...')
			results = wikipedia.summary(query)
			speak("According to Wikipedia")
			print(results)
			speak(results)

		# tested
		elif 'close' in query:
			speak('Have a good day. Bye')
			break
		
		# tested
		elif 'open youtube' in query:
			speak("What you want to search on youtube")
			command = takeCommand()
			query = ''
			for com in command.split():
				query += '+'+ com
			url = 'https://www.youtube.com/results?search_query=' + query
			webbrowser.open(url)
			speak('I am refering you to the youtube. And closing this application. Bye')
			break

		#  tested
		elif 'open google' in query:
			speak("What you want to search on google")
			command = takeCommand()
			query = ''
			for com in command.split():
				query += '+'+ com
			url = 'https://www.google.com/search?q=' + query
			webbrowser.open(url)
			# speak('I am refering you to the youtube. And closing this application. Bye')
			# break
		
		# pending
		# https://api.stackexchange.com/search/advanced?site=stackoverflow.com&q=firebase
		elif 'open stack overflow' in query:
			speak("What you want to search on stackoverflow")
			command = takeCommand()
			query = ''
			for com in command.split():
				query += '+'+ com
			url = 'https://stackoverflow.com/search?q=' + query + '&sort=relevance&s=c6d9666f-c256-4464-848f-9d96f91f1ba4'
			webbrowser.open(url)

		
		# tested
		elif 'play music' in query or "play song" in query:
			speak("Here you go with music")
			# music_dir = "G:\\Song"
			music_dir = r"C:\Users\Muhammad Ali Murtaza\Downloads\Video"
			songs = list(file for file in os.listdir(music_dir) if os.path.isfile(os.path.join(music_dir, file)))
			random = os.startfile(os.path.join(music_dir, songs[random. randint(0, len(songs))]))
			break

		# tested
		elif 'time' in query:
			strTime = datetime.datetime.now().strftime("%H:%M:%S") 
			strTime = datetime.datetime.now().strftime("%I:%M%p") 
			print(strTime)
			speak(f"The time is {strTime}")

		# tested
		elif 'email to richa' in query:
			try:
				speak("What should be the subject of Email")
				Subject = takeCommand()
				speak("What should be the body of Email")
				body = takeCommand()

				# Enter mail id of your friend here
				to_email = "richasinhacps@gmail.com"
				speak('Sending email at {0}'.format(to_email))
				sendEmail(subject=Subject, body=body, to_email=to_email)
				speak("Email has been sent !")
			except Exception as e:
				print(e)
				speak("I am not able to send this email")

		# tested
		elif 'send a mail' in query:
			try:
				speak("What should be the subject of Email")
				Subject = takeCommand()
				speak("What should be the body of Email")
				body = takeCommand()
				speak("whome should i send")
				to_email = input('Enter reciever email address: ')
				speak('Sending email at {0}'.format(to_email))
				sendEmail(subject=Subject, body=body, to_email=to_email)
				speak("Email has been sent !")
			except Exception as e:
				print(e)
				speak("I am not able to send this email")

		# tested
		elif 'how are you' in query:
			speak("I am fine, Thank you")
			speak("How are you, Sir")

		# tested
		elif 'fine' in query or "good" in query:
			speak("It's good to know that your fine")

		# tested
		elif "change my name to" in query:
			query = query.replace("change my name to", "")
			speak('Your name changed to: {0}'.format(query))
			assname = query

		# tested
		elif "change name" in query:
			speak("What would you like to call me, Sir ")
			assname = takeCommand()
			speak("Thanks for naming me")

		# tested
		elif "what's your name" in query or "What is your name" in query:
			speak("My friends call me")
			speak(assname)
			print("My friends call me", assname)

		# tested
		elif 'exit' in query or 'close' in query:
			speak("Thanks for giving me your time")
			exit()

		# tested
		elif "who made you" in query or "who created you" in query: 
			speak("I have been created by Muhammad Ali Murtaza.")
			
		# tested
		elif 'joke' in query:
			speak('Let me make a joke for you. Hmmm')
			joke = pyjokes.get_joke()
			print(joke) 
			speak(joke)

		# tested
		elif "weather" in query:
			
			# Google Open weather website
			# to get API of Open weather 
			api_key = "e2b6d54e8f7c27fc12a2f8ae4632a1fe"
			base_url = "http://api.openweathermap.org/data/2.5/weather?"
			speak(" City name ")
			print("City name : ")
			city_name = takeCommand()
			print("City name : ", city_name)
			complete_url = base_url + "appid=" + api_key + "&q=" + city_name
			response = requests.get(complete_url) 
			x = response.json() 
			
			if x["cod"] != "404": 
				y = x["main"] 
				current_temperature = y["temp"] 
				current_pressure = y["pressure"] 
				current_humidiy = y["humidity"] 
				z = x["weather"] 
				weather_description = z[0]["description"] 
				print(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidiy) +"\n description = " +str(weather_description)) 
				speak(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidiy) +"\n description = " +str(weather_description)) 
			
			else: 
				speak(" City Not Found ")
			

		# tested
		elif 'search' in query or 'play' in query:
			
			query = query.replace("search", "") 
			query = query.replace("play", "")		 
			webbrowser.open(query)

		# tested
		elif "who i am" in query:
			speak("If you talk then definitely your human.")

		# tested
		elif "why you came to world" in query:
			speak("Thanks to Gaurav. further It's a secret")

		# elif "calculate" in query: 
			
		# 	app_id = "Wolframalpha api id"
		# 	client = wolframalpha.Client(app_id)
		# 	indx = query.lower().split().index('calculate') 
		# 	query = query.split()[indx + 1:] 
		# 	res = client.query(' '.join(query)) 
		# 	answer = next(res.results).text
		# 	print("The answer is " + answer) 
		# 	speak("The answer is " + answer) 

		 

		# elif 'power point presentation' in query:
		# 	speak("opening Power Point presentation")
		# 	power = r"C:\\Users\\GAURAV\\Desktop\\Minor Project\\Presentation\\Voice Assistant.pptx"
		# 	os.startfile(power)

		elif 'is love' in query:
			speak("It is 7th sense that destroy all other senses")

		elif "who are you" in query:
			speak("I am your virtual assistant created by Gaurav")

		elif 'reason for you' in query:
			speak("I was created as a Minor project by Mister Gaurav ")

		# elif 'change background' in query:
		# 	ctypes.windll.user32.SystemParametersInfoW(20, 
		# 											0, 
		# 											"Location of wallpaper",
		# 											0)
		# 	speak("Background changed successfully")

		# elif 'open bluestack' in query:
		# 	appli = r"C:\\ProgramData\\BlueStacks\\Client\\Bluestacks.exe"
		# 	os.startfile(appli)
		
		# tested
		elif 'news' in query:
			
			try: 
				jsonObj = urlopen('''https://newsapi.org/v1/articles?source=the-times-of-india&sortBy=top&apiKey=99a6305c2291496e8e6451dc657b0d89''')
				data = json.load(jsonObj)
				i = 1
				
				speak('here are some top news from the times of india')
				print('''=============== TIMES OF INDIA ============'''+ '\n')
				
				for item in data['articles']:
					
					print(str(i) + '. ' + item['title'] + '\n')
					print(item['description'] + '\n')
					speak(str(i) + '. ' + item['title'] + '\n')
					i += 1
			except Exception as e:
				
				print(str(e))

		# tested
		elif 'lock window' in query:
				speak("locking the device")
				ctypes.windll.user32.LockWorkStation()

		
		elif 'shutdown system' in query:
				speak("Hold On a Sec ! Your system is on its way to shut down")
				subprocess.call('shutdown / p /f')

		elif 'empty recycle bin' in query:
			winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
			speak("Recycle Bin Recycled")

		# tested
		elif "don't listen" in query or "stop listening" in query:
			speak("for how much time you want to stop jarvis from listening commands")
			a = int(takeCommand())
			time.sleep(a)
			print(a)

		# tested
		elif "where is" in query:
			query = query.replace("where is", "")
			location = query
			speak("User asked to Locate")
			speak(location)
			webbrowser.open("https://www.google.nl/maps/place/" + location + "")

		# tested
		elif "camera" in query or "take a photo" in query:
			ec.capture(0, "Jarvis Camera ", "img.jpg")

		# restart
		elif "restart" in query:
			subprocess.call(["shutdown", "/r"])
			
		elif "hibernate" in query or "sleep" in query:
			speak("Hibernating")
			subprocess.call("shutdown / h")

		elif "log off" in query or "sign out" in query:
			speak("Make sure all the application are closed before sign-out")
			time.sleep(2)
			subprocess.call(["shutdown", "/l"])

		# tested
		elif "write a note" in query:
			speak("What should i write, sir")
			note = takeCommand()
			file = open('jarvis.txt', 'w')
			speak("Sir, Should i include date and time")
			snfm = takeCommand()
			if 'yes' in snfm or 'sure' in snfm:
				strTime = datetime.datetime.now().strftime("% H:% M:% S")
				file.write(strTime)
				file.write(" :- ")
				file.write(note)
			else:
				file.write(note)
		# tested
		elif "show note" in query:
			speak("Showing Notes")
			file = open("jarvis.txt", "r") 
			print(file.read())
			speak(file.read(5))

		elif "update assistant" in query:
			speak("After downloading file please replace this file with the downloaded one")
			url = '# url after uploading file'
			r = requests.get(url, stream = True)
			
			with open("Voice.py", "wb") as Pypdf:
				
				total_length = int(r.headers.get('content-length'))
				
				for ch in progress.bar(r.iter_content(chunk_size = 2391975),
									expected_size =(total_length / 1024) + 1):
					if ch:
					    Pypdf.write(ch)
		
		elif "jarvis" in query:
			
			wishMe()
			speak("Jarvis 1 point o in your service Mister")
			speak(assname)

		
			

		# 
		elif "wikipedia" in query:
			webbrowser.open("wikipedia.com")
		
		# tested 
		elif "good morning" in query:
			speak("A warm" +query)
			speak("How are you Mister")
			speak(assname)


		# tested
		# most asked question from google Assistant
		elif "will you be my gf" in query or "will you be my bf" in query: 
			speak("I'm not sure about, may be you should give me some time")

		# tested
		elif "how are you" in query:
			speak("I'm fine, glad you me that")

		# tested
		elif "i love you" in query:
			speak("It's hard to understand")


		# pending
		elif "what is" in query or "who is" in query:
			
			# Use the same API key 
			# that we have generated earlier
			client = wolframalpha.Client("QJY2RY-W3X23XLEVV")
			res = client.query(query)
			
			try:
				print (next(res.results).text)
				speak (next(res.results).text)
			except StopIteration:
				print ("No results")

		else:
			speak('I am unable to understand your command. Please repeat')
			# Command go here
			# For adding more commands

	
    

