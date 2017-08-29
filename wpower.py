from tkinter import *
import time
import datetime
from openpyxl import load_workbook
from apscheduler.schedulers.blocking import BlockingScheduler
import pyttsx3 as vol

# volume related issues change hear later needs to eb worked on##############
def pronounce_meaning(mean):
    engine=vol.init()
    engine.setProperty('rate',200)
    engine.setProperty('volume',0.7)
    engine.setProperty('voice',1)
    engine.setProperty('language',"en")    
    engine.say('means')
    time.sleep(1)    
    engine.say(mean)
    engine.runAndWait()
    engine.stop()
def pronounce(word,mean):
    engine=vol.init()
    engine.setProperty('rate',200)
    engine.setProperty('volume',0.7)
    #engine.setProperty('languages','fr')
    engine.say(word)
    engine.runAndWait() 
    engine.stop()
    #engine.voice.getProperty('languages')
    time.sleep(2)
    #engine.setProperty('languages','en-us') 
    pronounce_meaning(mean)
    
    
    
    # or stop() can be used improve this later
# ###############################################################################    
def update(wordlist):
    global root
    for key,value in wordlist.items():
        var1.set(key)
        var2.set(value)
        mainwindow.update()
        pronounce(key,value)
        time.sleep(10)
    root.destroy() 
# GUI interface needs to e improved ########################################
# improve design ###############
class App:
         
    def __init__(self, master):
        global mainwindow     
        mainwindow = Frame(master)
        global var1
        global var2
        var1 = StringVar()
        label1 = Label( master,compound=CENTER,width=20,border=0,height=2,font = "Helvetica 20 bold",
                        bg="orange",fg="dark green", textvariable=var1,
                        relief=RAISED )
        
        var1.set("The word")
        label1.pack()  
        var2 = StringVar()
        label2 = Label( master,compound=CENTER,width=20,height=2,border=0,font = "Helvetica 20 bold",
                        bg="orange",fg="dark green", textvariable=var2,
                        relief=RAISED )
                
        var2.set("Meaning ")
        label2.pack()         
        mainwindow.pack(fill=BOTH, expand=YES)
        
    def getWords():
        wbook=load_workbook('./wordList.xlsx')
        sheetList=wbook.get_sheet_names()
        sheet=wbook.get_sheet_by_name(sheetList[0])
            
        wordList={}
        for i in range(1,20):
            if sheet.cell(row=i,column=1).value!=None:
                wordList['{}'.format(sheet.cell(row=i,column=1).value)]='{}'.format(sheet.cell(row=i,column=2).value)
                    
            else:break 
        return wordList               
# ############################# #########################################   
def placement(win):
    win.update_idletasks()
    width = win.winfo_width()
    height = win.winfo_height()
    win.geometry('{}x{}+{}+{}'.format(width,height, win.winfo_screenwidth()-width,win.winfo_screenheight()-height))          

# ######################################################################
# main code
def main():
    global root
    root = Tk()
    root.attributes('-alpha',0.5)
    root.wm_attributes('-topmost','true')
    root.wm_attributes('-disabled','true')
    root.overrideredirect(True)
    #root.title("Pack - Example 12")
    display = App(root)
    placement(root)
    wordlist=App.getWords()
    update(wordlist)
    root.mainloop()
    
# schedule to appear on the ottom right of the screen   ###################
def schedule():
    scheduler = BlockingScheduler()
    scheduler.add_job(main, 'interval', minutes=5)
    scheduler.start()
    
# #########################################################################    
schedule()
    
""" needs to be done 
1 waiting time etween appearance on the screen

2 word selection

3 randomisation

4 counting and deletion after a number of appearance

5 voice on and off checkox

6 change language between languages

7 process design from installation to uninstallation

"""