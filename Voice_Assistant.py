import requests, json, webbrowser, sys, os, speech_recognition, wmi, pathlib, pandas, openpyxl, xlrd
from bs4 import BeautifulSoup
import datetime as date
from io import StringIO
import geocoder as gc
import AppOpener as op
from geopy.geocoders import Nominatim
from neuralintents import GenericAssistant
import pyttsx3 as tts
import datetime as dt

#defines did not understand string
dnu = "I did not understand you, please try again"

#making the speech recogniser and speaker
recog = speech_recognition.Recognizer()
voice = tts.init()
voice.setProperty('rate', 150)
voice.setProperty('voice', 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_EN-US_ZIRA_11.0')

#to hold the responses
#this is not done in the json file, because we also want to display the message in text to the user
responses = []

#function for getting input from the user
def get_input():
    global recog
    done = False
    while not done:
        try:
            with speech_recognition.Microphone() as mic:
                recog.adjust_for_ambient_noise(mic, duration = 0.2)
                output = recog.listen(mic)

                output = recog.recognize_google(output)
                output = output.lower()
                done = True
            
        except speech_recognition.UnknownValueError:
            recog = speech_recognition.Recognizer()
            voice.say(dnu)
            voice.runAndWait()
    return output
            
def create_note():
    global recog
    #defining the voice responses for this function
    responses.append("What should the note be?")
    responses.append("What would you like the file to be called?")
    responses.append("The note was created under the specified name")

    #getss what user wants as a note
    voice.say(responses[0])
    voice.runAndWait()

    output= get_input() + "\n"
    #gets what user wants the note name to be
    voice.say(responses[1])
    voice.runAndWait()
    
    fname = get_input()
    fname = fname.lower()
    noteDate = date.datetime.today()
    noteDate = date.datetime.strftime(noteDate, '%Y-%m-%d')
    fname+="_" + str(noteDate)
    #writes the note
    with open(f'{fname}.txt', 'w') as f:
        f.write(output + "\n")
        voice.say(responses[2])
        voice.runAndWait()
    responses.clear()

def add_toDo():
    global recog
    #defining the specific responses
    responses.append("What would you like to add to your to do list?")
    responses.append("The item was added to your to do list")

    voice.say(responses[0])
    voice.runAndWait()

    output = get_input()
    
    with open('toDo.txt', 'a') as f:
        f.write(output)
        done = True
        voice.say(responses[1])
        voice.runAndWait()
    responses.clear()

def show_toDo():
    global recog
    responses.append("Sure thing! Here is your to do list")
    voice.say(responses[0])
    f = open('toDo.txt', 'r')
    for x in f:
        voice.say(x)

    voice.runAndWait()
    responses.clear()

def get_weather():
    global recog
    #defines the responses

    responses.append("Here is today's weather:")

    geolocator = Nominatim(user_agent = 'geoapiExercises')
    voice.say(responses[0])
    
    latlng = [str(gc.ip('me').lat), str(gc.ip('me').lng)]
    location = geolocator.reverse(latlng[0]+","+latlng[1])
    location = (location.raw['address']).get('city')
    url = "https://api.openweathermap.org/data/2.5/weather?"
    api = 'dc48750e6e63794686d487cdd950ef62'
    url = url + 'q='+location+'&appid='+api + '&units=metric'
    response = requests.get(url)
    if(response.status_code==200):
        #getting the first bit of data
        data = response.json()
        main = data['main']
        descMain = data['weather']
        winds = data['wind']

        #transforms the data

        temp = str(round(float(main['temp'])))
        description = descMain[0]['description']
        wind_speed = winds['speed']
        deg = winds['deg']
        wind_direction = '.'
        if(deg > 0 and deg< 90):
            wind_direction = 'North-East'
        elif deg >90 and deg < 180:
            wind_direction = 'South-East'
        elif deg > 180 and deg < 270:
            wind_direction = 'South-West'
        elif deg > 270 and deg < 360:
            wind_direction = 'North-West'
        elif deg == 0 or deg == 360:
            wind_direction = 'North'
        elif deg == 90:
            wind_direction = 'East'
        elif deg == 180:
            wind_direction = 'South'
        elif deg == 270:
            wind_direction = 'West'

        #gets the final output of the speaker
        responses.append("The current temperature is " + temp + "degrees Celcius")
        responses.append("with " + description)
        responses.append("Wind speeds are " + str(wind_speed) + " kilometers per hour " + wind_direction)

        for i in range(1,len(responses)):
            voice.say(responses[i])
        voice.runAndWait()

        responses.clear()

        """
def make_timetable():
    global recog
    #defines responses
    responses.append("Sure thing! But the format will be specific")
    responses.append("Please say your module codes as follows:")
    responses.append("CMPG 311 AND MTHS 311 AND NONE AND NONE AND MTHS312")
    responses.append("The NONE is when you have an empty timeslot")
    days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
    for i in range(0,4):
        responses.append("Wha do you have on " + days[i])
    #makes the excel file required
    noList = [None,None,None,None,None]
    df = pandas.DataFrame({"Day":days,"07:30":noList, "09:30":noList, "11:00":noList, "13:00":noList, "14:30":noList, "16:00":noList})
    df.to_excel('timetable.xlsx', index=False, header=True)
    
    voice.say(responses[0])
    voice.say(responses[1])
    voice.say(responses[2])
    voice.say(responses[3])
    voice.runAndWait()

    #gets the required input
    final = [[], [], [], [], []]
    for i in range(0,len(days)):
        voice.say("What do you have on " + days[i])
        voice.runAndWait()
        output = get_input()
        #formats input
        output = output.replace("and", ",")
        output = output.replace(" ","")
        output = output.split()
        print(output)
        final[i] = (output)
    temp = []
    wb = openpyxl.load_workbook('timetable.xlsx')
    ws=wb.active
    
    for i in range(0,6):
        temp.append(chr(66+i))

    for i in range(0,len(temp)):
        for j in range(0,len(days)):
            finalTemp = temp[i] + str(j+2)
            ws[finalTemp] = final[j][i]

    wb.save('timetable.xlsx')
    responses.clear()
    """

def see_timetable():
    responses.append("Sorry, but it doesn't seem that a timetable exists.")
    responses.append("A file was given to you called timetables. Please fill it in as specified")
    responses.append("Please add an ")
    responses.append("   EN   ")
    responses.append("When you do not have a module")
    
    responses.append("It is a weekend, so you have no modules!")
    responses.append("Sure thing!")
    path = str(pathlib.Path(__file__).parent.resolve())
    path+="\\timetable.xlsx"
    if not os.path.isfile(path):
        voice.say(responses[0])
        days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
        noList = [None,None,None,None,None]
        df = pandas.DataFrame({"Day":days,"07:30":noList, "09:30":noList, "11:00":noList, "13:00":noList, "14:30":noList, "16:00":noList})
        df.to_excel('timetable.xlsx', index=False, header=True)
        for i in range(1, 5):
            voice.say(responses[i])
        voice.runAndWait()
    else:
        today = dt.datetime.today().weekday()
        #0 is monday and so forth
        match today:
            case 0:
                today = 'Monday'
            case 1:
                today = 'Tuesday'
            case 2:
                today = 'Wednesday'
            case 3:
                today = 'Thursday'
            case 4:
                today = 'Friday'
            case _:
                today = 'Monday'

        if today == 'not_valid':
            voice.say(responses[5])
            voice.runAndWait()
        else:
            df = pandas.read_excel(path)
            
            row = df[df['Day'] == today]
            col = (row!= 'N').any()
            row = row.loc[:, col]
            row = row.drop(['Day'], axis=1)

            vals = []
            for column in row:
                col = row[column]
                vals.append(str(col.values) + " at " + str(column))
            
            final=''
            for x in vals:
                final+=' ' + x
            
            final = final.replace('[\'', '')
            final = final.replace('\']', '')
            
            voice.say(responses[6])
            voice.say('Today you have ' + final)
            voice.runAndWait()

def hello():
    responses.append("Hi")
    responses.append("What can I do for you")
    
    voice.say(responses[0])
    voice.say(responses[1])
    voice.runAndWait()

    responses.clear()

def google_search():
    global recog
    responses.append("Sure thing! What would you like to search?")
    voice.say(responses[0])
    voice.runAndWait()

    search = get_input()
                
    url = "https://www.google.com.tr/search?q=" + search
    webbrowser.open_new_tab(url)
    voice.runAndWait()
    responses.clear()

def exit_program():
    responses.append("Okay, have a wonderful day!")
    voice.say(responses[0])
    voice.runAndWait()
    responses.clear()
    sys.exit(0)

def open_efundi():
    url = "https://casprd.nwu.ac.za/cas/login?service=https%3A%2F%2Fefundi.nwu.ac.za%2Fsakai-login-tool%2Fcontainer"
    webbrowser.open_new_tab(url)

def open_app():
    global recog
    responses.append("Sure thing, what would you like to open?")
    voice.say(responses[0])
    voice.runAndWait()
    output = get_input()
    op.open(output, match_closest=True)
    responses.clear()

def open_email():
    responses.append("Sure thing!")
    voice.say(responses[0])
    voice.runAndWait()
    responses.clear()
    webbrowser.open_new_tab("gmail.com")
    

def get_loadshedding():
    url = "https://www.ourpower.co.za/areas/jb-marks/potchefstroom?block=5"
    page = requests.get(url)

    soup = BeautifulSoup(page.content, "html.parser")
    results = soup.find("div", class_="home_nextOffMsg__8MJvu")
    loadshedding=''
    loadshedding += results.text.strip()

    voice.say(loadshedding)
    voice.runAndWait()


mappings = {
    "greeting": hello,
    "create_notes": create_note,
    "add_to_do_list": add_toDo,
    "show_to_do_list": show_toDo,
    "weather": get_weather,
    'search': google_search,
    'open_application': open_app,
    'see_timetable': see_timetable,
    'open_efundi': open_efundi,
    'open_email': open_email,
    'show_loadshedding': get_loadshedding
}



assistant = GenericAssistant('intents.json', intent_methods=mappings)
assistant.train_model()
assistant.save_model("voice_assistant")
assistant.load_model("voice_assistant")

while True:
    try:
        print("I await instruction")
        voice.runAndWait()
        with speech_recognition.Microphone() as mic:
            recog.adjust_for_ambient_noise(mic, duration=0.2)
            audio = recog.listen(mic)

            message = recog.recognize_google(audio)
            message = message.lower()
            print(message)
            assistant.request(message)
    except speech_recognition.UnknownValueError:
        recog = speech_recognition.Recognizer()
