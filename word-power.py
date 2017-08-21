from tkinter import *

class App:
    def __init__(self, master):
        mainwindow = Frame(master)
             
        var = StringVar()
        label1 = Label( root,compound=CENTER,width=25,font = "Helvetica 16 bold",bg="light grey",fg="dark green", textvariable=var, relief=RAISED )
        
        var.set("The word")
        label1.pack()  
        var = StringVar()
        label2 = Label( root,compound=CENTER,width=25,font = "Helvetica 16 bold",bg="light grey",fg="dark green", textvariable=var, relief=RAISED )
                
        var.set("Meaning ")
        label2.pack()         
        """word = Text(mainwindow)
        word.insert(INSERT, "this will be the word")
        word.insert(END, ".")
        
        meaning=Text(mainwindow)
        meaning.insert(INSERT,"this will be the meaning")
        word.pack()

        meaning.pack()"""
        mainwindow.pack(fill=BOTH, expand=YES)
def center(win):
    win.update_idletasks()
    width = win.winfo_width()
    height = win.winfo_height()
    win.geometry('{}x{}+{}+{}'.format(width, height, win.winfo_screenwidth()-width, 0))        
root = Tk()
root.attributes('-alpha',0.5)
root.wm_attributes('-topmost','true')
root.wm_attributes('-disabled','true')
root.overrideredirect(True)
#root.title("Pack - Example 12")
display = App(root)
center(root)
root.mainloop()