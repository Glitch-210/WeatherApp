#import packages
import requests
import json
import win32com.client as wincom
import time

#enter city name
city = input("Enter your city name:")
url=f"https://api.weatherapi.com/v1/current.json?key=8b61fc85e6f24dc7a7694427243011&q={city}"

r= requests.get(url) #this will request to the url
wdict = json.loads(r.text) #it loads string

#speak variable
speak = wincom.Dispatch("SAPI.SpVoice")

#data collection
temp=wdict['current']['temp_c']
region= wdict['location']['region']
country = wdict['location']['country'] 
lp = wdict['current']['last_updated'] 
ws = (wdict['current']['wind_kph']) 
wd = (wdict['current']['wind_degree'] )  

def WeatherApp():        
    while True:
        choice = int(input("What do you want to know about the city?\n 1.Temprature\n 2.Region\n 3.Country\n 4.last updated time\n 5.wind speed(kph)\n 6.wind degree\n 7.ALL\n 8.Exit\n "))

        match choice:
            case 1:
                text = f"Temperature of {city} is {temp}"
                print(text)
                speak.Speak(text)
            case 2:
                text = f"Region of {city} is {region}"
                print(text)
                speak.Speak(text)
            case 3: 
                text = f"Country of {city} is {country}"
                print(text)
                speak.Speak(text)
            case 4: 
                text = f"Last updated time is {lp}"
                print(text)
                speak.Speak(text)
            case 5: 
                text = f"Wind Speed of {city} is {ws} kph"
                print(text)
                speak.Speak(text)
            case 6: 
                text = f"Wind degree of {city} is {wd} degree"
                print(text)
                speak.Speak(text)
            case 7: 
                text = f"Temperature of {city} is {temp}"
                speak.Speak(text)
                time.sleep(1)
                text = f"Region of {city} is {region}"
                speak.Speak(text)
                time.sleep(1)
                text = f"Country of {city} is {country}"
                speak.Speak(text)
                time.sleep(1)
                text = f"Last updated time is {lp}"
                speak.Speak(text)
                time.sleep(1)
                text = f"Wind Speed of {city} is {ws} kph"
                speak.Speak(text)
                time.sleep(1)
                text = f"Wind degree of {city} is {wd} degree"
                speak.Speak(text)
                time.sleep(1)
            case 8:
                return False
    
WeatherApp()