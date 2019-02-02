"""
GUI with 2 tabs.
Tab 1: Search Outlook e-mails based on filters. Must specify which folders to search before running.
Tab 2: Connect to SQL database for quick item lookups and quickly access customer portal
"""



from tkinter import *
from tkinter import scrolledtext
import io
from datetime import datetime, timedelta
from tkcalendar import Calendar
from tkinter.scrolledtext import ScrolledText
import weblogin
from tkinter import ttk

#I understand globals should be avoided, but nonetheless they served OK here.
notred=True
global date
date = (datetime.now() + timedelta(days=-30)).strftime("%y-%m-%d %H:%M")


global commit
commit = None

""""
Save Folder Range Button: Outlook directory sequence changes with each startup, so it is necessary tell Python to scan all folders and find the corresponding directory tree
"""
def savefolderrange():
    if foldertolookin.get():
        foldertolookin1 = (foldertolookin.get())
        foldertolookin1 =foldertolookin1.split(',')
        from emailstructure import getemailstructure, saveresults
        index = getemailstructure(foldertolookin1)
        saveresults('valuefile' ,index)
        

    global commit
    commit = 1


def clicked():

    weblogin.accessmylam(selected, accountvalue,partno.get(),name_password1, name_password2 )

#get the name and password for each account from the text file and save them as list sets
def getpasswordsets():
    passwordsets = []
    with open("login.txt") as data:
        datalines = (line.rstrip('\r\n').replace(' ', '') for line in data)
        for line in (datalines):

            passwordsets.append(line.split(","))


    
    return passwordsets[0], passwordsets[1]

name_password1, name_password2 =getpasswordsets()
print (name_password1, name_password2 )
#GUI
main = Tk()
window = main
nb = ttk.Notebook(main)

# adding Frames as pages for the ttk.Notebook 
page1 = ttk.Frame(nb)

page1.grid( sticky = N+S+W+E)
page1.columnconfigure(0, weight=1)
page1.rowconfigure(2, weight=1)
# second page
page2 = ttk.Frame(nb)

page2.grid( sticky = N+S+W+E)
page2.columnconfigure(0, weight=1)
page2.rowconfigure(2, weight=1)

page3 = ttk.Frame(nb)

page3.grid( sticky = N+S+W+E)
page3.columnconfigure(0, weight=1)
page3.rowconfigure(2, weight=1)
nb.add(page1, text='E-mail Tools')
nb.add(page2, text='Part Number Search')

nb.grid(sticky = N+S+W+E)
        
#Execute Search of Outlook e-mails after the folder range to look in has been specified
def gogo():
    editArea.delete('1.0', END)
    
    buffer = io.StringIO()
    editArea.insert(END,"processing \n")
    from emailer import findstringinemail

    if commit ==1 and foldertolookin.get() =="":
        foldertolookin1 = None
    else:
        text1 = open('valuefile.txt', 'r')
        foldertolookin1 = text1.read()
        text1.close()
        foldertolookin1 = eval(foldertolookin1)
    print(searchstring.get(), foldertolookin1,desiredsender.get(),attachments.get(),allchains.get(), date)
    
    editArea.tag_config('target', foreground="red")

    if searchstring.get():
        if foldertolookin1 != None :
            for index in range(len(foldertolookin1)):
                buffer = findstringinemail(searchstring.get(), foldertolookin1[index],desiredsender.get(),attachments.get(),allchains.get(), date)
                output = buffer.getvalue()
                
                char_list = [output[j] for j in range(len(output)) if ord(output[j]) in range(65536)]
                output=''
                for j in char_list:
                    output = output+ j
                
                
                editArea.insert(END, output)
                
        if foldertolookin1 == None :
            buffer = findstringinemail(searchstring.get(), None,desiredsender.get(), attachments.get(),allchains.get(), date)
            output = buffer.getvalue()
            char_list = [output[j] for j in range(len(output)) if ord(output[j]) in range(65536)]
            output=''
            for j in char_list:
                output = output+ j
            editArea.insert(END, output)
            #text.see(END)
    editArea.insert(END,"done")
    

#Configuring the GUI to my personal liking
window.columnconfigure(0, weight=1)
window.rowconfigure(0, weight=1)
main.iconbitmap(r'pachi1.ico')

window.title("E-mail Searcher")


window.geometry('900x800')
w = window.winfo_reqwidth()
h = window.winfo_reqheight()
ws = window.winfo_screenwidth()
hs = window.winfo_screenheight()
x = (ws/2) - (w/2)
y = (hs/2) - (h/2)
window.geometry('+%d+%d' % (x, y)) 


#creation of menu elements
menubar = Menu(main)


 # create a pulldown menu, and add it to the menu bar
filemenu = Menu(menubar, tearoff=0)
srchStr = "test"
filemenu.add_separator()
filemenu.add_command(label="Exit", command=main.quit)
menubar.add_cascade(label="File", menu=filemenu)



helpmenu = Menu(menubar, tearoff=0)
helpmenu.add_command(label="version 2018.09.04 by Wesley Chu")
menubar.add_cascade(label="About", menu=helpmenu)



# display the menu
main.config(menu=menubar)
foldertolookin = StringVar()


#GUI for all the filters the user can set when performing an e-mail search
Label(page1, text= "Folders to Search in: ").grid(row=0, column=0,sticky = N+E)
Entry(page1, textvariable = foldertolookin, width = 75).grid(row=0, column = 1, columnspan = 4,sticky = N+W)
savebutton = Button(page1, command=savefolderrange, padx=10, text="Save Folder Range \n (Leave Blank to Recall)")
savebutton .grid(column=4, row=0,columnspan = 1, sticky=N+S+E+W)


attachments = BooleanVar()
attachments.set(False)
Label(page1, text= "Attachments Only").grid(row=0, column=5,sticky = N+E)
rad1 = Radiobutton(page1,text='On' ,value=True, variable=attachments, tristatevalue=1)

rad2 = Radiobutton(page1,text='Off', value=False, variable=attachments, tristatevalue=0)

rad1.grid(column=6, row=0, sticky = N+W)

rad2.grid(column=7, row=0, sticky = N+W)


Label(page1, text= "Full E-mail Chains").grid(row=1, column=5,sticky = N+E)

allchains = BooleanVar()
allchains.set(False)

allchains1 = Radiobutton(page1,text="On",value=True, variable=allchains, tristatevalue=1)

allchains2 = Radiobutton(page1,text="Off", value=False, variable=allchains, tristatevalue=0)

allchains1.grid(column=6, row=1, sticky = N+W)


allchains2.grid(column=7, row=1, sticky = N+W)




searchstring = StringVar()
desiredsender = StringVar()
Label(page1, text= "Search For: ").grid(row=1, column=0,sticky = N+E)
Entry(page1, textvariable = searchstring).grid(row=1, column=1,sticky = N+W)
Label(page1, text= "Sender: ").grid(row=1, column=2,sticky = N+E)
Entry(page1, textvariable =desiredsender).grid(row=1, column=3,sticky = N+E)



go = Button(page1, command=gogo, padx=10, text="GO")
go.grid(column=8, row=0,rowspan = 2, sticky=N+S+E+W)




#date calendar pop-up so user can specify how far back to search e-mails in
def date1():
    def print_sel():
        cal.destroy()
        top.destroy()
        print(cal.selection_get())
        global date
        date =  cal.selection_get().strftime("%y-%m-%d %H:%M")
        top.destroy()
    top = Toplevel(window)
    date = datetime.now()
    cal = Calendar(top
                   , selectmode='day',
                    year=datetime.today().year, month=datetime.today().month, day=datetime.today().day)
    cal.pack(fill="both", expand=True)
    Button(top, text="OK", command=print_sel).pack()

Button(page1, text='Input Date', command=date1).grid(row=1, column=4,sticky = N+E)


#creating a blank space that is scrollable. This will be where e-mail results are displayed
editArea = ScrolledText(
master = page1,
wrap   = WORD,
height = 40    
)

editArea.grid(row = 2, column = 0, columnspan = 12, sticky = N+S+W+E)
editArea.columnconfigure(0, weight=1)
editArea.rowconfigure(2, weight=1)

#Highlight the keyword. In hindsight this is probably not needed and word should always be highlighted for easy reading
def highlight():
    if searchstring.get() != "":
        global notred
        editArea.tag_config("black", foreground="black")
        editArea.tag_config("red", foreground="red") 
        editArea.tag_remove("black", "1.0", "end")
        editArea.tag_remove("red", "1.0", "end")
    
        count = IntVar()
        editArea.mark_set("matchStart", "1.0")
        editArea.mark_set("matchEnd", "1.0")
        print(notred)
        while True:
            index = editArea.search(searchstring.get(), "matchEnd","end", count=count)
            if index == "": break # no match was found
    
            editArea.mark_set("matchStart", index)
            editArea.mark_set("matchEnd", "%s+%sc" % (index, count.get()))
    
            if notred ==True:
                editArea.tag_add("red", "matchStart", "matchEnd")
                
            else:
                editArea.tag_add("black", "matchStart", "matchEnd")
        notred = not notred
    else:
        return
    
    
        
Button(page1, text='Highlight', command=highlight).grid(row=1, column=4,sticky = N+W)





# second tab starts here





#Customer website has 2 different login settings, which can be specified through this section
accountvalue = IntVar(None, 2)
accountvalue.set(1)

account1 = Radiobutton(page2,text=name_password1[0],value=1, variable=accountvalue, tristatevalue=1)

account2 = Radiobutton(page2,text=name_password2[0], value=2, variable=accountvalue, tristatevalue=0)

account1.grid(row=0, column=5,sticky = N+E)

partno = StringVar()
account2.grid(row=0, column=6,sticky = N+E)
Label(page2, text= "Part No.").grid(row=0, column=7,sticky = N+E)
Entry(page2, textvariable = partno).grid(row=0, column=8,sticky = N+E)
partfound = StringVar()


#when a part number is searched for, this calls the function and displays it in the display section
def partinfocaller():
    text.delete('1.0', END)
    buffer1 = io.StringIO()
    buffer1 = getpartinfo(partno)
    print(buffer1)
    output = buffer1.getvalue()
    print(output)
    text.insert(END, output) 
    

#Interface for second tab
from mySQL import getpartinfo
go = Button(page2, padx=1, command= partinfocaller, text="Part Look-up")
go.grid(row = 0, column = 9)


selected = IntVar(None, 2)
selected.set(1)

rad1 = Radiobutton(page2,text='No Action' ,value=1, variable=selected, tristatevalue=1)

rad2 = Radiobutton(page2,text='get Current Rev', value=2, variable=selected, tristatevalue=0)

rad3 = Radiobutton(page2,text='get BOM', value=3, variable=selected, tristatevalue=0)

rad4 = Radiobutton(page2,text='Current Rev + BOM', value=4, variable=selected, tristatevalue=0)

rad1.grid(row=1, column=5,sticky = N+E)

rad2.grid(row=1, column=6,sticky = N+E)
rad3.grid(row=1, column=7,sticky = N+E)

rad4.grid(row=1, column=8,sticky = N+E)
gotosite = Button(page2, padx=1, command=clicked, text="Go To Website")
gotosite.grid(row = 1, column = 9)

text = ScrolledText(page2)
text.grid(row = 2, column = 0, columnspan = 12, sticky = N+E+W+S)
text.columnconfigure(0, weight=1)
text.rowconfigure(2, weight=1)


#wanted to add info from an excel log file that tracks daily tasks but didnt find it worthwhile to implement yet
"""

def geteventlog():
    text3.delete('1.0', END)
    from eventlog import eventlog2
    datastore = eventlog2()
    text3.insert(END, datastore) 

#third tab starts here
from html import *
go = Button(page3, padx=1, command= App, text="HTML")
go.grid(row = 0, column = 9)
go = Button(page3, padx=1, command= geteventlog, text="Get Event Log")
go.grid(row = 1, column = 9)
text3 = ScrolledText(page3)
text3.grid(row = 2, column = 0, columnspan = 12, sticky = N+E+W+S)
text3.columnconfigure(0, weight=1)
text3.rowconfigure(2, weight=1)
"""
window.mainloop()


