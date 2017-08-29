import pyttsx3
engine = pyttsx3.init()
engine.say("il y a deux voiture")
engine.setProperty('rate',190)  #120 words per minute
engine.setProperty('volume',0.9) 
engine.runAndWait()