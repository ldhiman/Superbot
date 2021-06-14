import wolframalpha
import pyttsx3
import tkinter
from newsapi import NewsApiClient
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
from pygame import mixer
import shutil
from twilio.rest import Client
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
from selenium import webdriver
import pyautogui

#define voice
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

#define speak function
def speak(audio):
	engine.say(audio)
	engine.runAndWait()

#wish on  startup
def wishMe():
	hour = int(datetime.datetime.now().hour)
	if hour >= 0 and hour < 12:
		speak("Good Morning Sir !")

	elif hour >= 12 and hour < 16:
		speak("Good Afternoon Sir !")

	else:
		speak("Good Evening Sir !")

	speak("I am your Assistant")
	speak('Superbot.')

#Find any file
def find_files(filename, search_path):
	response = os.popen("wmic logicaldisk get caption")
	list1 = []
	total_file = []
	for line in response.readlines():
		line = line.strip("\n")
		line = line.strip("\r")
		line = line.strip(" ")
		if (line == "Caption" or line == ""):
			continue
		list1.append(line)
	return list1

#Listen function
def listen():
	r = sr.Recognizer()

	with sr.Microphone() as source:
		print("Listening...")
		r.pause_threshold = 1
		audio = r.listen(source, phrase_time_limit=4)
	return audio

#take the user command
def takeCommand():
	if command_type == 'audio':
		r = sr.Recognizer()
		audio = listen()
		try:
			print("Recognizing...")
			query = r.recognize_google(audio, language='en-in')
			print(f"User said: {query}\n")
			startup_sound()

		except Exception as e:
			print(e)
			print("Unable to Recognize your voice.")
			speak('Sorry, Speak that again')
			return 'None'
		return query

	else:
		query = input('Enter the Command :')
		return query

#send email
def send_email(user, domain, sub, content):
	email = f"{user}@{domain}"
	print(email)
	try:
		os.system("taskkill /im chrome.exe /f")
	except Exception:
		pass

	options = webdriver.ChromeOptions()
	options.add_argument("user-data-dir=C:\\Users\\91935\\AppData\\Local\\Google\\Chrome\\User Data")
	driver = webdriver.Chrome(options=options)

	driver.get('https://accounts.google.com/b/0/AddMailService')

	driver.find_element_by_xpath('//*[@id=":jm"]/div/div').click()

	pyautogui.sleep(5)

	def find_thumb():
		abc = r'C:\Users\91935\Pictures\subject.png'
		sent = r'C:\Users\91935\Pictures\sent.png'
		abc = pyautogui.locateCenterOnScreen(abc)
		if abc is not None:
			pyautogui.moveTo(abc)
			pyautogui.click()
			pyautogui.write(sub)
			pyautogui.move(0, -35)
			pyautogui.click()
			pyautogui.write(email)
			pyautogui.moveTo(abc)
			pyautogui.move(0, 35)
			pyautogui.click()
			pyautogui.write(content)
			sent = pyautogui.locateCenterOnScreen(sent)
			pyautogui.moveTo(sent)
			pyautogui.click()
			print("email sent!!!!")
		else:
			find_thumb()
	find_thumb()

#startup_sound
def startup_sound():
	mixer.init()
	mixer.music.load("Siri_Sound_Effect_HD.mp3")
	mixer.music.play()
	time.sleep(1)
	mixer.music.stop()

#method of conservation
def method():
	speak("Which method you want to prefer Audio or text??")
	ans = input('Enter your answer??? 1. for Audio. 2. for Text  (answer 1/2)')
	ans = int(ans)
	if ans == 1 or 2:
		if ans == 1:
			command_type = 'audio'
			speak('Method choose Audio')
			return command_type
		else:
			command_type = 'text'
			speak('Method choose Audio')
			return command_type
	else:
		print("Wrong Input!! Please give a correct input.")

#functioning of the command given by user
def function():
	query = takeCommand().lower()

	#For wikipedia
	if 'wikipedia' in query:
		try:
			speak('Searching Wikipedia...')
			query = query.replace("wikipedia", "")
			results = wikipedia.summary(query, sentences=3)
			speak("According to Wikipedia")
			print(results)
			speak(results)
		except Exception:
			speak("Sorry for in convience unable to sreaching wikipedia")

	#for youtube
	elif 'open youtube' in query:
		try:
			speak("Here you go to Youtube\n")
			webbrowser.open("www.youtube.com")
		except Exception:
			speak("Sorry, Unable to open youtube")

	#for sreaching query on google
	elif 'open google' in query:
		try:
			speak("Here you go to Google\n")
			speak('what you want to search??')
			search = takeCommand()
			search = search.replace(" ", "+")
			webbrowser.open(f'https://www.google.com/search?q={search}')
		except Exception:
			speak("Unable to open google")

	#open stackoverflow		
	elif 'open stackoverflow' in query:
		try:
			speak("Here you go to Stack Over flow")
			webbrowser.open("www.stackoverflow.com")
		except Exception:
			speak("Unable to open stackoverflow")

	elif 'open new tap' in query:
		webbrowser.open_new_tab("chrome://newtab")

	#play music
	elif 'play music' in query or "play song" in query:
		try:
			options = webdriver.ChromeOptions()
			options.add_argument("--start-maximized")
			driver = webdriver.Chrome(options=options)
			speak("Here you go with music")
			driver.get('https://music.youtube.com/')
			speak("which music you want to listen??")
			music = takeCommand()
			driver.find_element_by_xpath('//*[@id="icon"]').click()
			pyautogui.write(music)
			pyautogui.press('enter')

			def find_thumb():
				play = r'C:\Users\91935\Pictures\top.png'
				play = pyautogui.locateCenterOnScreen(play)
				if play is not None:
					pyautogui.moveTo(play)
					pyautogui.move(0, 60)
					pyautogui.click()
				else:
					find_thumb()

			find_thumb()
		except Exception:
			speak("Sorry, Unable to play song!!")

	elif 'virtual mouse mode' in query:
		speak('Activating Virtual Mouse Mode')
		time.sleep(0.2)
		speak('Virtual Mouse Mode Activated')
		import AiVirtualMouseProject

	#Switch the song
	elif 'play another music' in query or 'play another song' in query:
		try:
			options = webdriver.ChromeOptions()
			options.add_argument("--start-maximized")
			driver = webdriver.Chrome(options=options)
			speak("which music you want to listen??")
			music = takeCommand()
			driver.find_element_by_xpath('//*[@id="icon"]').click()
			pyautogui.write(music)
			pyautogui.press('enter')

			def find_thumb():
				play = r'C:\Users\91935\Pictures\top.png'
				play = pyautogui.locateCenterOnScreen(play)
				if play is not None:
					pyautogui.moveTo(play)
					pyautogui.move(0, 60)
					pyautogui.click()
				else:
					find_thumb()

			find_thumb()
		except Exception:
			speak("Unable to change song")

	elif 'type' in query:
		speak('What are the keywords')
		keywords = takeCommand()
		pyautogui.write(keywords)

	elif 'space' in query:
		pyautogui.press('space')

	elif 'enter' in query:
		pyautogui.press('enter')		

	#current time
	elif 'what is the time' in query:
		try:
			strTime = datetime.datetime.now().strftime("%I:%M %p")
			speak(f"Sir, the time is {strTime}")
		except Exception:
			speak('Sorry,Unable to find time')

	#open opera
	elif 'open browser' in query:
		webbrowser.open('https://www.google.com/')

	elif 'open chrome' in query:
		webbrowser.open('https://www.google.com/')	
			
	elif 'send a mail' in query:
		try:
			speak('what is the username??')
			user = takeCommand()
			speak("what is the domain name (example: gmail.com, yahoo.com, outlook.com)")
			domain = takeCommand()
			speak("What is the subject, if any??")
			sub = takeCommand()
			speak("What is the content")
			content = takeCommand()
			send_email(user, domain, sub, content)
			speak("Email has been sent !")
		except Exception as e:
			speak("I am not able to send this email")

	elif 'how are you' in query:
		speak("I am fine, Thank you...")
		speak("How are you, Sir")

	elif 'fine' in query or "good" in query:
		speak("It's good to know that your fine....")
		speak("What should i do for you???")

	elif "change my name to" in query:
		speak("I am unable to change your name...")
		speak("Function could not be installed")

	elif "change your name" in query:
		speak("Sorry, but i have already have a name Siri")
		speak("is that good name???")

	elif "what's your name" in query or "What is your name" in query:
		speak("My friends call me, Siri!!")
		print("My friends call me,Siri")

	elif 'exit' in query or 'goodbye' in query or 'bye' in query:
		speak("Thanks for giving me your time")
		exit()

	elif "who made you" in query or "who created you" in query:
		speak("I have been created by Lucky.")
		speak('for performing your minor tasks')

	elif 'joke' in query:
		try:
			jokes = pyjokes.get_joke()
			speak(jokes)
			print(jokes)
		except Exception:
			speak('Unable to perform task, Sorry')

	elif "calculate" in query:
		try:
			app_id = "PJYEUX-VQ25TRP6J5"
			client = wolframalpha.Client(app_id)
			indx = query.lower().split().index('calculate')
			query = query.split()[indx + 1:]
			res = client.query(' '.join(query))
			answer = next(res.results).text
			print("The answer is " + answer)
			speak("The answer is " + answer)
		except Exception:
			speak("Unable to calculate")

	elif 'search' in query or 'play news' not in query and 'play' in query:

		query = query.replace("search", "")
		query = query.replace("play", "")
		webbrowser.open(f'https://www.google.com/search?q={query}')

	elif "who i am" in query:
		speak("If you talk then definately your human.")

	elif "why you came to world" in query:
		speak("Thanks to Mr.Lucky. further It's a secret")

	elif 'power point presentation' in query:
		speak("opening Power Point presentation")
		try:
			power = find_files('powerpoint','C:')
			os.startfile(power)
		except Exception:
			speak("File not found")

	elif 'what is love' in query:
		speak("It is 7th sense that destroy all other senses")

	elif "who are you" in query:
		speak("I am your Superbot created by Lucky")

	elif 'reason for you' in query:
		speak("I was created for preforming Minor task given by you!!")

	elif 'play news' in query:
		try:
			jsonObj = urlopen(
				'https://newsapi.org/v1/articles?source=the-times-of-india&sortBy=top&apiKey=962cb59479a64b76955224be79b56f60')
			data = json.load(jsonObj)
			i = 1



			speak('here are some top news from the times of india')
			print('''=============== TIMES OF INDIA ============''' + '\n')
			for item in data['articles']:
				print(str(i) + '. ' + item['title'] + '\n')
				print(item['description'] + '\n')
				speak(str(i) + '. ' + item['title'] + '\n')
				i += 1


		except Exception as e:

			print(str(e))

	elif 'lock window' in query or 'sign out' in query:
		try:
			speak("locking the device")
			ctypes.windll.user32.LockWorkStation()
		except Exception:
			speak('Sorry, unable to lock window')

	elif 'shutdown system' in query or 'shutdown pc' in query or 'shutdown computer' in query:
		try:
			import subprocess
			speak("Hold On a Sec ! Your system is on its way to shut down")
			subprocess.call('shutdown / p /f')
			exit()
		except Exception:
			speak('Sorry, unable to shutdown system')

	elif 'empty recycle bin' in query:
		try:
			winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
			speak("Recycle Bin Recycled")
		except Exception:
			speak('Sorry, unable to perform action')

	elif "don't listen" in query or "stop listening" in query:
		try:
			speak("for how much time you want to stop from listening commands")
			a = takeCommand()
			if 'minute' in a:
				a.replace('minute', '')
				a = int(a)
				a = a*60
				time.sleep(a)
			elif 'hour' in a:
				a.remove('hour', '')
				a = int(a)
				a = a*60*60
				time.sleep(a)
			elif 'second' in a:
				a.remove('second', '')
				a = int(a)
				time.sleep(a)
			else:
				speak('unexcepted duration..')
		except Exception:
			speak('Sorry, unable to perform task due to some error')

	elif "where is" in query:
		try:
			query = query.replace("where is", "")
			location = query
			speak("User asked to Locate")
			speak(location)
			webbrowser.open("https://www.google.nl/maps/place/" + location + "")
		except Exception:
			speak('Sorry, unable to find place..')

	elif "camera" in query or "take a photo" in query:
		try:
			ec.capture(0, "Camera", "img.jpg")
		except Exception:
			speak('Unable to perform task')

	elif "restart the pc" in query:
		try:
			import subprocess
			subprocess.call(["shutdown", "/r"])
			exit()
		except Exception:
			speak('Unable to restart')

	elif "hibernate" in query or "sleep" in query:
		try:
			speak("Hibernating")
			import subprocess
			subprocess.call("shutdown / h")
		except Exception:
			speak('Sorry,Unable to sleep')

	elif "log off" in query or "sign out" in query:
		try:
			speak("Make sure all the application are closed before sign-out")
			time.sleep(5)
			import subprocess
			subprocess.call(["shutdown", "/l"])
		except Exception:
			speak('Unable to perform task')

	elif "write a note" in query:
		try:
			speak("What should i write, sir")
			note = takeCommand()
			file = open('Note.txt', 'w') 
			speak("Sir, Should i include date and time")
			snfm = takeCommand()
			if 'yes' in snfm or 'sure' in snfm:
				strTime = datetime.datetime.now().strftime("% H:% M:% S")
				file.write(strTime)
				file.write(" :- ")
				file.write(note)
			else:
				file.write(note)
		except Exception:
			speak('Unable to make note...')

	elif "show note" in query:
		try:
			speak("Showing Notes")
			file = open("Note.txt", "r")
			print(file.read())
			speak(file.read(6))
		except Exception:
			speak('Unbale to perform task')

	elif "weather" in query:

		# Google Open weather website
		# to get API of Open weather
		api_key = "1f2378e55a0a5c8c8467d76ad7e43717"
		base_url = "http://api.openweathermap.org/data/2.5/weather?"
		speak(" City name ")
		print("City name : ")
		city_name = takeCommand()
		complete_url = base_url + "appid =" + api_key + "&q =" + city_name
		response = requests.get(complete_url)
		x = response.json()

		if x["cod"] != "404":
			y = x["main"]
			current_temperature = y["temp"]
			current_pressure = y["pressure"]
			current_humidiy = y["humidity"]
			z = x["weather"]
			weather_description = z[0]["description"]
			print(" Temperature (in kelvin unit) = " + str(
				current_temperature) + "\n atmospheric pressure (in hPa unit) =" + str(
				current_pressure) + "\n humidity (in percentage) = " + str(
				current_humidiy) + "\n description = " + str(weather_description))

		else:
			speak(" City Not Found ")

	elif "send message " in query:
		# You need to create an account on Twilio to use this service
		account_sid = 'Account Sid key'
		auth_token = 'Auth token'
		client = Client(account_sid, auth_token)

		message = client.messages \
			.create(
			body=takeCommand(),
			from_='Sender No',
			to='Receiver No'
		)

		print(message.sid)

	elif "wikipedia" in query:
		webbrowser.open("wikipedia.com")

	elif "Good Morning" in query:
		speak("A warm" + query)
		speak("How are you Mister Sir")


	# most asked question from google Assistant
	elif "will you be my gf" in query or "will you be my bf" in query:
		speak("I'm not sure about, may be you should give me some time")

	elif "how are you" in query:
		speak("I'm fine, glad you me that")

	elif "i love you" in query:
		speak("It's hard to understand")

	elif "what is" in query or "who is" in query:

		# Use the same API key
		# that we have generated earlier
		client = wolframalpha.Client("API_ID")
		res = client.query(query)

		try:
			print(next(res.results).text)
			speak(next(res.results).text)
		except StopIteration:
			print("No results")

#main file
if __name__ == '__main__':
	clear = lambda: os.system('cls')
	clear()
	wishMe()

	while True:
		command_type = 'audio'
		function()
		startup_sound()

