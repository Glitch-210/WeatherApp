import requests
import json
import win32com.client as wincom
city = input("Enter your city name:")
url=f"https://api.weatherapi.com/v1/current.json?key=8b61fc85e6f24dc7a7694427243011&q={city}"

r= requests.get(url)
# print(r.text)
wdict = json.loads(r.text) #it loads string
w = (wdict['current']['temp_c'])

speak = wincom.Dispatch("SAPI.SpVoice")
text = f"The current temperature in {city} is {w} degrees"
speak.Speak(text)