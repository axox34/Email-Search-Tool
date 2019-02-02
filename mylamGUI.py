"""
GUI for accessing customer website portal and fetches information on parts
deprecated, use emailsearchGUI
"""


from tkinter import *
import weblogin
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

#GUI
main = Tk()
window = main
main.iconbitmap(r'C:\Users\wchu\Documents\LiClipse 5.0.1\Workspace\pdfing\dva.ico')


window.title("myLAM Quick Tool")


window.geometry('500x125')
w=500
h = 100
ws=window.winfo_screenwidth()
hs=window.winfo_screenheight()
x=(ws/2)-(w/2)
y=(hs/2)-(h/2)
window.geometry('+%d+%d'%(x,.5*y))


#creation of menu elements
menubar = Menu(main)


 # create a pulldown menu, and add it to the menu bar
filemenu = Menu(menubar, tearoff=0)

filemenu.add_separator()
filemenu.add_command(label="Exit", command=main.quit)
menubar.add_cascade(label="File", menu=filemenu)


helpmenu = Menu(menubar, tearoff=0)
helpmenu.add_command(label="version 2018.08.03")
menubar.add_cascade(label="Help", menu=helpmenu)



# display the menu
main.config(menu=menubar)


#GUI for selecting the appropriate formula file
selectframe=Frame(main).grid(row=0, columnspan=2)
selectvalue = IntVar(None, 2)
selectvalue.set(1)

select1 = Radiobutton(selectframe,text="MyLam",value=1, variable=selectvalue, tristatevalue=1)

select2 = Radiobutton(selectframe,text="Amat", value=2, variable=selectvalue, tristatevalue=0)
select1.grid(column=0, row=0, sticky = N+W)
select2.grid(column=1, row=0, sticky = N+W+E)





radiobutton_frame0 = Frame(main, width = 300, height=30).grid(row=1, columnspan=4)



accountvalue = IntVar(None, 2)
accountvalue.set(1)

account1 = Radiobutton(radiobutton_frame0,text=name_password1[0],value=1, variable=accountvalue, tristatevalue=1)

account2 = Radiobutton(radiobutton_frame0,text=name_password2[0], value=2, variable=accountvalue, tristatevalue=0)

account1.grid(column=0, row=1, sticky = N+W+E)

partno = StringVar()
account2.grid(column=1, row=1, sticky = N+W+E)
Label(radiobutton_frame0, text= "Part No.").grid(row=1, column=2,sticky = N+E)
Entry(radiobutton_frame0, textvariable = partno).grid(row=1, column=3,sticky = N+W+E)


#GUI for selecting the appropriate formula file



radiobutton_frame = Frame(main).grid(row=1, columnspan=4)


selected = IntVar(None, 2)
selected.set(1)

rad1 = Radiobutton(radiobutton_frame,text='No Action' ,value=1, variable=selected, tristatevalue=1)

rad2 = Radiobutton(radiobutton_frame,text='get Current Rev', value=2, variable=selected, tristatevalue=0)

rad3 = Radiobutton(radiobutton_frame,text='get BOM', value=3, variable=selected, tristatevalue=0)

rad4 = Radiobutton(radiobutton_frame,text='Current Rev + BOM', value=4, variable=selected, tristatevalue=0)

rad1.grid(column=0, row=2, sticky = N+W)

rad2.grid(column=1, row=2, sticky = N+W)

rad3.grid(column=2, row=2, sticky =N+W, padx= 5)

rad4.grid(column=3, row=2, sticky = N+W, padx= 5)





go_frame = Frame(main, width = 50, height=60).grid(row=1, column=4)
go = Button(go_frame, command=clicked, padx=10, text="GO")
go.grid(column=4, row=1,rowspan = 2, sticky=N+S+E+W)






window.mainloop()




