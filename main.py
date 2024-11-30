#import packages
import requests
import json
import win32com.client as wincom
import time

#enter city name
city = input("Enter your city name:")
city = city.capitalize()
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
        choice = int(input(f"What do you want to know about the {city}?\n 1.Temprature\n 2.Region\n 3.Country\n 4.Last Updated Time\n 5.Wind speed(kph)\n 6.Wind degree\n 7.ALL\n 8.Exit\n "))

        match choice:
            case 1:
                text = f"Temperature of {city} is {temp} degree celcius"
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
                text = f"Temperature of {city} is {temp} degree celcius"
                print(text)
                speak.Speak(text)
                time.sleep(0.5)

                text = f"Region of {city} is {region}"
                print(text)
                speak.Speak(text)
                time.sleep(0.5)

                text = f"Country of {city} is {country}"
                print(text)
                speak.Speak(text)
                time.sleep(0.5)

                text = f"Last updated time is {lp}"
                print(text)
                speak.Speak(text)
                time.sleep(0.5)
                
                text = f"Wind Speed of {city} is {ws} kph"
                print(text)
                speak.Speak(text)
                time.sleep(0.5)
                
                text = f"Wind degree of {city} is {wd} degree"
                print(text)
                speak.Speak(text)
                time.sleep(0.5)
            case 8:
                return False
    
WeatherApp()