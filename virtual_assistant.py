# Descriptipn: This is a virtual assistant program that gets the date, current time, responds back with a random
#              greeting, returns information on a person with wikipedia, open some application in computer i.e. Chrome
#              Mozilla, Visual Studio, Word, Excel, NotePad and open some URL address YouTube, Facebook etc.

#pip install SpeechRecognition - for recognizing the voice command and converting to text
#pip install gTTS - Google Text To Speech, for converting the given text to speech
#pip install wikipedia
#pip install requests
#pip install wolframalpha - for calculation given by user
#pip install selenium - for web based work from browser

# Import the libraries

import speech_recognition as sr
import os
from gtts import gTTS
import datetime
import warnings
import calendar
import random
import wikipedia
import requests
import wolframalpha
import pyaudio
import webbrowser

# Ignore any warning messages
warnings.filterwarnings('ignore')

SearchLetter = ["search", "looking", "looking for", "look something", "go google", "Go", "go", "open", "Open"]
ApplicationLetter = ["Play", "Start", "play"]
GeneralLetter = ["date", "time", "who", "how", "what"]

# Record audio and return it as a string
def recordAudio():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print('Say Something!')
        audio = r.listen(source)

# Use Googles speech recognition
    data = ''
    try:
        data = r.recognize_google(audio)
        print('you said: '+data)
    except sr.UnknownValueError:
        print('Google Speech Recognition could not understand the audio, unknown error')
    except sr.RequestError as e:
        print('Request results from Google Speech Recognition service error '+ e)

    return data

# A Function to get the virtual assistant response
def assistantResponse(text):

    print(text)
    # Convert the text to speech
    myobj = gTTS(text= text, lang='en', slow= False)

    # Save the converted audio to a file
    myobj.save('assistant_response.mp3')

    #Plat the converted file
    os.system('start assistant_response.mp3')

# A function for wake word(s) or phrase
def wakeWord(text):
    WAKE_WORDS = ['hey computer', 'okey computer'] # A list of wake words
    text = text.lower() # Converting the text all lower case words

    #Chech to see if the users command/text contains a wake word/phrase
    for phrase in WAKE_WORDS:
        if phrase in text:
            return True

    # If the wake word isn't found in the text from the loop and so it return False
    return False

# A function to get the current date
def getDate():

    now = datetime.datetime.now()
    my_date = datetime.datetime.today()
    weekday = calendar.day_name[my_date.weekday()] # For example Friday
    monthNum = now.month
    dayNum = now.day

    # A list of months
    month_names = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September',
                   'October', 'November', 'December']

    # A list of ordinal numbers
    ordinalNumbers= ['1st', '2nd', '3rd', '4th', '5th', '6th', '7th', '8th', '9th', '10th', '11th', '12th',
                     '13th', '14th', '15th', '16th', '17th', '18th', '19th', '20th', '21st', '22nd', '23rd',
                     '24th', '25th', '26th', '27th', '28th', '29th', '30th', '31st', ]

    return 'Today is ' +weekday+' '+month_names[monthNum-1]+' the '+ordinalNumbers[dayNum-1]+'. '

# A function to return a random greeting response
def greeting(text):

    # Greeting inputs
    GREETING_INPUTS = ['hi', 'hey', 'hola', 'greetings', 'whats up', 'hello']

    #Greeting responses
    GREETING_RESPONSES = ['howdy', 'whats good', 'hello', 'hey there']

    # If the users input is a greeting, then return a randomly chosen greeting response
    for word in text.split():
        if word.lower() in GREETING_INPUTS:
            return random.choice(GREETING_RESPONSES) +'.'

    #If no greeting was detected then return an empty string
    return ''

# A function to get a persons first and last name from the text
def getPerson(text):

    wordList = text.split() #Splitting the text into a list of words

    for i in range(0, len(wordList)):
        if i + 3 <= len(wordList) - 1 and wordList[i].lower() == 'who' and wordList[i+1].lower() == 'is':
            return wordList[i+2] + ' '+ wordList[i+3]

def getComputer(text):

    wordList2 = text.split()

    for i in range(0, len(wordList2)):
        if i + 3 <= len(wordList2) - 1 and wordList2[i].lower() == 'how' and wordList2[i+1].lower() == 'are':
            return wordList2[i+2] + ' '+ wordList2[i+3]

# A function to get a open Chrome application
def getOpenApplication(text):

    word = text.split()
    chrome_path = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'
    driver = webbrowser.get(chrome_path)
    driver.open('https://www.google.com/?#q=' + word[3])

def getWordApplication(text):

    os.startfile('C:/Program Files/Microsoft Office/root/Office16/WINWORD.EXE')

def getExcelApplication(text):

    os.startfile('C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE')

def getVsApplication(text):

    os.startfile('C:/Program Files (x86)/Microsoft Visual Studio/2017/Community/Common7/IDE/devenv.exe')

def getUnityApplication(text):

    os.startfile('C:/Program Files/Unity Hub/Unity Hub.exe')

def getNotePadApplication(text):

    os.startfile('C:/Users/nevza/AppData/Roaming/Microsoft/Windows/Start Menu/Programs/Accessories/Notepad.lnk')

def getOpenFirefoxApplication(text):

    firefox_path = 'C:/Program Files/Mozilla Firefox/firefox.exe %s'
    driver = webbrowser.get(firefox_path)
    driver.open('https://www.google.com/')

while True:

    # Record the Audio
    text = recordAudio()
    response = ''

    # Check for the wake word/phrase
    if(wakeWord(text) == True):

        # Check for greetings by the user
        response = response + greeting(text)

        # Check to see if the user said anything having to do with the date
        if('date' in text):
            get_date = getDate()
            response = response + ' '+get_date

        # Check to see if the user said anything having to do with te time
        if('time' in text):
            now = datetime.datetime.now()
            meridiem =''
            if now.hour >=12:
                meridiem = 'p.m' # Post Meridiem (PM) after midday
                hour = now.hour - 12
            else:
                meridiem = 'a.m' # Ante Meridiem (AM) before midday
                hour = now.hour

            # Convert minute into a proper string
            if now.minute < 10:
                minute = '0'+str(now.minute)
            else:
                minute = str(now.minute)

            response = response +' '+'It is '+str(hour)+ ':'+ minute+ ' '+meridiem+' .'

        # Check to see if the user said 'who is'
        if('who is' in text):
            person = getPerson(text)
            wiki = wikipedia.summary(person, sentences=2)
            response = response +' '+ wiki

        # Check to see if the user said 'how are you' computer answered
        if('how are you' in text):
            computer = getComputer(text)
            response = response +' '+'I am well'+' .'

        # User said 'who are you' computer introduces itself
        if('who are'in text):
            response = response +' '+'I am a computer. Nevzat who designed me. ' \
                                     'And its designed to help you with your applications and calls related ' \
                                     'to your voice commands'+' .'

        # User said 'what are you doing'
        if('what are you doing' in text):
            response = response +' '+'Listening to you'+' .'

        # User said 'I want to say something'
        if('I want to say something' in text):
            response = response +' '+'Yes, Listening to you'+' .'

        # If user says 'open Chrome' Computer opening Chrome and response
        if('Chrome' in text):
            response = response +' '+'Opening Chrome Application'+' .'
            chrome = getOpenApplication(text)

        # If user says 'open Firefox' computer opening firefox application and response answer
        if('Firefox' in text):
            firefox = getOpenFirefoxApplication(text)
            response = response +' '+'Opening Firefox Application'+' .'

        # If users want to be open youtube, only say 'hey computer open YouTube' and computer will open this app
        if('YouTube' in text):
            youtube = getOpenApplication(text)
            response = response +' '+'Opening Youtube Application'+' .'

        # Like this youtube application. Only say 'hey computer open Facebook' and computer will open this app
        if('Facebook' in text):
            facebook = getOpenApplication(text)
            response = response +' '+'Opening Facebook Application'+' .'

        # User says 'open Microsoft Word' Computer opening Microsoft Word
        if('Microsoft Word' in text):
            response = response +' '+'Opening Microsoft Word'+' .'
            word = getWordApplication(text)

        # User says 'open Microsoft Excel' Computer opening Microsoft Excel
        if('Microsoft Excel' in text):
            response = response +' '+'Opening Microsoft Excel'+' .'
            excel = getExcelApplication(text)

        # User says 'open Visual Studio' Computer opening Visual Studio
        if('Visual Studio' in text):
            response = response +' '+'Opening Visual Studio'+' .'
            vs = getVsApplication(text)

        # User says 'open Unity Hub' Computer opening Unity Hub
        if('Unity hub' in text):
            response = response +' '+'Opening Unity Hub'+' .'
            unity = getUnityApplication(text)

        # User says 'open Notepad' Computer opening Notepad
        if('Notepad' in text):
            response = response +' '+'Opening Note Pad'+' .'
            notepad = getNotePadApplication(text)

        # Have the assistant respond back using audio and the text from response
        assistantResponse(response)