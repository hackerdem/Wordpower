from tkinter import *
import time
import random
import datetime
from openpyxl import load_workbook
from apscheduler.schedulers.blocking import BlockingScheduler
import pyttsx3 as vol
import os,sys
from unidecode import unidecode as udec
import threading
global volume_value
def fixWord(word):
    # this section added sice pronounciation is not good for french words containing signs
    
    return udec(word)
def placement(win):
    win.update_idletasks()
    width = win.winfo_width()
    height = win.winfo_height()
    win.geometry('{}x{}+{}+{}'.format(width,height, win.winfo_screenwidth()-width,win.winfo_screenheight()-height))          

    
def get_words():
    
    global sheet
    global repoSheet
    global discardedWords
    global wbook
    global sheetList
    
    wbook=load_workbook('./wordList.xlsx')
    try:
        
        
        sheetList=wbook.get_sheet_names()
        sheet=wbook.get_sheet_by_name(sheetList[0])
        repoSheet=wbook.get_sheet_by_name(sheetList[1])
        discardedWords=wbook.get_sheet_by_name(sheetList[2])
        print(sheet.max_row)
        wordList={}
        global size
        size=sheet.max_row
        random_start=random.randrange(1,size,1)
        #print(random_start)
        for i in range(1,51):
            iterator=random_start+i if random_start+i<=size else random_start+i+1-size
            if sheet.cell(row=iterator,column=1).value!=None:
                try:
                    wordList['{}'.format(sheet.cell(row=iterator,column=1).value)]=[sheet.cell(row=iterator,column=2).value,sheet.cell(row=iterator,column=3).value]
                except Exception:pass
            else:
                break 
        #print(wordList)
        update(wordList)
    except UnicodeDecodeError:
        get_words()
    #return wordList  
def pronounce_meaning(mean):
    #print(volume_value.get())
    engine1=vol.init()
    engine1.setProperty('rate',180)
    engine1.setProperty('volume',volume_value.get())
    v=engine1.getProperty('voices')
    engine1.setProperty('voice',v[1].id)   
    engine1.say('means')
    time.sleep(1)    
    engine1.say(mean)
    engine1.runAndWait()
    engine1.stop()
def pronounce(word,mean):
    if volume_value.get()!='0.0':
        engine=vol.init()
        engine.setProperty('rate',200)
        engine.setProperty('volume',volume_value.get())
        engine.say(fixWord(word))
        engine.runAndWait() 
        engine.stop()
        time.sleep(1)
        pronounce_meaning(mean)
    else:pass
    
    
    

# ###############################################################################    
def update_excel(key,value):
    try:
        
        for i in range(1,size):
            
            if sheet.cell(row=i,column=1).value==key:
                
                sheet['C{}'.format(i)]=value[1]
            
                wbook.save('./wordList.xlsx')
            else:pass
    except Exception as e:
        #print(e)
        pass
def volume_on_off():
    if volume_value.get()!='0.0':
        volume_value.set('0.0') 
        volume_caption.set("Volume Off") 
    else:
        volume_value.set('0.7')
        volume_caption.set("Volume On")
def user_discard(fr,eng):
    #print(fr,eng)
    discard_word(fr,[eng,1])
    
def discard_word(key,value):
    
    val_list=[key,value[0],value[1]]
    try:
        for i in range(1,size):
            
            if sheet.cell(row=i,column=1).value==key:
                counter=size
                
                while repoSheet.cell(row=counter,column=1).value==None:
                    # we should know if there is no value in the repo
                    if counter==1: 
                        #print("There is no value in Repo")
                        break
                    
                    counter=counter-1
            
                discard=1
                while discardedWords.cell(row=discard,column=1).value!=None:
        
                    discard=discard+1
                columns=['A','B','C']
                for d in columns:
            
                    discardedWords['{}{}'.format(d,discard)]=val_list[columns.index(d)]
                    
                for i in range(1,size):
                    if key==sheet.cell(row=i,column=1).value:
                        for k in columns:
                            cval=1+columns.index(k)
                            sheet['{}{}'.format(k,i)]=None
                            sheet['{}{}'.format(k,i)]=repoSheet.cell(row=counter,column=cval).value
                            repoSheet['{}{}'.format(k,counter)]=None
                        break
                wbook.save('./wordList.xlsx') 
    except Exception as e:
        print(e)
        
       
def update(wordlist):
    
    try:
        for key,value in wordlist.items():
            if value[1]!=None and value[1]>99:# there is a problem here gives error sometimes
                discard_word(key,value)
            else:
                french_word.set(key.strip())
                english_word.set(value[0].strip())
                value[1]=1 if value[1]==None else value[1]+1
                root.update()
               
                update_excel(key, value)
                pronounce(key,value[0])
                time.sleep(10)
        
    except Exception as e:
        #print(e)
        pass
# GUI interface needs to e improved ########################################
# improve design ###############
def wait():
    os.system("pause")
def exit_from_application():
    wbook.save('./wordList.xlsx') 
    os._exit(1)
def main():

    PROGRAM_NAME='FRENCH WORD LEARNING'
    root=Tk()
    
    root.overrideredirect(True)
    root.configure(background='#118977')
    root.attributes('-alpha',0.8)
    root.wm_attributes('-topmost','true')
    root.title(PROGRAM_NAME)
    global english_word
    global french_word 
    global root
    global volume_value,volume_caption
    volume_value=StringVar()
    french_word=StringVar()
    english_word=StringVar()
    volume_caption=StringVar()
    volume_value.set('0.7')
    french_word.set(" ")
    english_word.set(" ")
    volume_caption.set("Volume On")
    root.resizable(width=False, height=False)
    french=Label(fg='black',textvariable=french_word,relief=GROOVE,width=64,height=3).grid(row=0,padx=1,pady=1,column=0,columnspan=3)
    
    english=Label(textvariable=english_word,relief=GROOVE,width=64,height=6).grid(row=1,padx=1,pady=1,column=0,columnspan=3)
    discard_button=Button(text='Discard',command=lambda : user_discard(french_word.get(),english_word.get()),relief=GROOVE,width=20,height=2).grid(row=2,padx=1,pady=1,column=0,sticky='e')
    voice_button=Button(textvariable=volume_caption,command=lambda :threading.Thread(target=volume_on_off).start(),relief=GROOVE,width=20,height=2).grid(row=2,padx=1,pady=1,column=1,sticky='e')
    exit_button=Button(text='Exit',command=lambda :threading.Thread(target=exit_from_application).start(),relief=GROOVE,width=20,height=2).grid(row=2,padx=1,pady=1,column=2,sticky='e')
    placement(root)
    w=threading.Thread(target=get_words).start()
    root.mainloop()
    
    
main()  
# schedule to appear on the bottom right of the screen   ###################
"""def schedule():
    scheduler = BlockingScheduler()
    scheduler.add_job(main, 'interval', minutes=5)
    scheduler.start()
    
# #########################################################################    
schedule()"""
        
""" needs to be done 
1 waiting time between appearance on the screen*************************

2 word selection

3 randomisation DONE!!!!!!!!!!!!!!!!!!

4 counting and deletion after a number of appearance-done************** DONE!!!!!!!!!!!!!!!!!! 100 times

5 voice on and off checkox DONE!!!!!!!!!!!!!!!!!!

6 change language between languages******************************* DONE!!!!!!!!!!!!!!!!!

7 process design from installation to uninstallation

"""