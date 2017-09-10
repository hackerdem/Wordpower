from tkinter import *
import time
import datetime
from openpyxl import load_workbook
#from openpyxl.utils import coordinate_from_string as rowcol
from apscheduler.schedulers.blocking import BlockingScheduler
import pyttsx3 as vol
import os
import pandas as pd
# volume related issues change hear later needs to eb worked on##############
def pronounce_meaning(mean):
    engine=vol.init()
    engine.setProperty('rate',200)
    engine.setProperty('volume',0.7)
    engine.setProperty('voice',1)
    engine.setProperty('language',"en-us")    
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
def update_excel(key,value):
    try:
        
        for i in range(1,200):
            
            if sheet.cell(row=i,column=1).value==key:
                
                sheet['C{}'.format(i)]=value[1]
            
                wbook.save('./wordList.xlsx')
            else:pass
    except Exception as e:
        print(e)
        pass
def discard_word(key,value):
    
    val_list=[key,value[0],value[1]]
    try:
        for i in range(1,200):
            
            if sheet.cell(row=i,column=1).value==key:
                counter=200
                
                while repoSheet.cell(row=counter,column=1).value==None:
                    # we should know if there is no value in the repo
                    if counter==1: 
                        print("There is no value in Repo")
                        break
                    
                    counter=counter-1
            
                discard=1
                while discardedWords.cell(row=discard,column=1).value!=None:
        
                    discard=discard+1
                columns=['A','B','C']
                for d in columns:
                    discardedWords['{}{}'.format(d,discard)]=val_list[columns.index(d)]
                    
                for i in range(1,200):
                    if key==sheet.cell(row=i,column=1).value:
                        for k in columns:
                            sheet['{}{}'.format(k,i)]=None
                            sheet['{}{}'.format(k,i)]=repoSheet.cell(row=counter,column=int(1+columns.index(k))).value
                            repoSheet['{}{}'.format(k,counter)]=None
                        break
                wbook.save('./wordList.xlsx') 
    except Exception as e:
        print(e)
        
       
def update(wordlist):
    global root
    try:
        for key,value in wordlist.items():
            if value[1]>99:
                discard_word(key,value)
            else:
                var1.set(key)
                var2.set(value[0])
                value[1]=1 if value[1]==None else value[1]+1
                update_excel(key, value)
                mainwindow.update()
                pronounce(key,value[0])
                time.sleep(10)
        root.destroy() 
    except Exception as e:
        print(e)
        pass
# GUI interface needs to e improved ########################################
# improve design ###############
def wait():
    os.system("pause")
def hidden():
    print('hello')
    """label1.width=0
    label1.height=0
    mainwindow.update()"""

class App:
    # make label and button a class####################
    def __init__(self, master):
        global mainwindow     
        mainwindow = Frame(master)
        global var1,var2
        global label1
        var1 = StringVar()
        label1 = Label( master,compound=CENTER,width=21,border=0,height=2,font = "Helvetica 20 bold",
                        bg="#FF9900",fg="#CC3333", textvariable=var1,
                        relief=RAISED )
        
        
        label1.pack()  
        var2 = StringVar()
        label2 = Label( master,compound=CENTER,width=21,height=2,border=0,font = "Helvetica 20 bold",
                        bg="#FF9900",fg="#CC3333", textvariable=var2,
                        relief=RAISED )
                
        
        label2.pack()  
        """self.pause = Button(master, 
                             text="PAUSE",bg='#CC3333',border=0, fg="#FFCC00",width=8,
                             command=wait)
        self.play = Button(master, 
                                     text="PLAY",bg='#CC3333',border=0, fg="#FFCC00",width=8,
                                     command=hidden)  
        self.hide = Button(master, 
                                             text="HIDE",bg='#CC3333',border=0, fg="#FFCC00",width=8,
                                             command=hidden)  
        self.hide.bind(hidden)
        self.discard = Button(master, 
                                             text="DISCARD",bg='#CC3333',border=0, fg="#FFCC00",width=8,
                                             command=lambda : discard_word(
                                                                          'a')    )
        self.voice = Button(master, 
                                                     text="VOICE",bg='#CC3333',border=0, fg="#FFCC00",width=8,
                                                     command=quit)          
        self.pause.pack(side=LEFT)
        self.play.pack(side=LEFT)
        self.hide.pack(side=LEFT)
        self.discard.pack(side=LEFT)
        self.voice.pack(side=LEFT)
        self.slogan = Button(master,width=8,bg='#CC3333',border=0, fg="#FFCC00",
                                 text="POSITION")
        self.slogan.pack(side=LEFT)"""        
        mainwindow.pack(fill=BOTH, expand=YES)
        
    def getWords():
        global sheet
        global repoSheet
        global discardedWords
        global wbook
        global sheetList
        wbook=load_workbook('./wordList.xlsx')
        
        sheetList=wbook.get_sheet_names()
        sheet=wbook.get_sheet_by_name(sheetList[0])
        repoSheet=wbook.get_sheet_by_name(sheetList[1])
        discardedWords=wbook.get_sheet_by_name(sheetList[2])
            
        wordList={}
        for i in range(1,20):
            if sheet.cell(row=i,column=1).value!=None:
                try:
                    wordList['{}'.format(sheet.cell(row=i,column=1).value)]=[sheet.cell(row=i,column=2).value,sheet.cell(row=i,column=3).value]
                except Exception:pass
            else:
                
                for key,value in wordList.items():
                    """try:print(key,value[0])
                    except Exception:pass"""
                break 
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
    #root.wm_attributes('-disabled','true') # when enabled no interaction with ui
    root.overrideredirect(True)
    #root.title("Pack - Example 12")
    display = App(root)
    placement(root)
    wordlist=App.getWords()
    update(wordlist)
    root.mainloop()
main()  
"""# schedule to appear on the ottom right of the screen   ###################
def schedule():
    scheduler = BlockingScheduler()
    scheduler.add_job(main, 'interval', seconds=5)
    scheduler.start()
    
# #########################################################################    
schedule()"""
    
""" needs to be done 
1 waiting time etween appearance on the screen

2 word selection

3 randomisation

4 counting and deletion after a number of appearance

5 voice on and off checkox

6 change language between languages

7 process design from installation to uninstallation

"""