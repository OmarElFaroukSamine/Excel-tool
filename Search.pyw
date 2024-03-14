#Imports
import openpyxl
import tkinter as tk
import tkinter.messagebox 


# Designing window 
root = tk.Tk()
root.config(background='white')
root.eval('tk::PlaceWindow . center')
window_height = 200
window_width = 400
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x_cordinate = int((screen_width/2) - (window_width/2))
y_cordinate = int((screen_height/2) - (window_height/2))
root.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))
root.title("Searching tool")
root.iconbitmap('Data/Icon.ico')
# Declaring string variables
fn=tk.StringVar()
ln=tk.StringVar()
mn=tk.StringVar()


# Getting the variables
def svtos():
    firstName=fn.get()
    middleName=mn.get()
    lastName=ln.get()
    fullname = firstName+middleName+lastName
    fullname = fullname.upper()
    fullname = fullname.replace(" ", "")

#Check the input has numbers
def cn(string):
    return any(char.isdigit() for char in string)


#The script that searches the file    
def search():
    #Getting the values from input fields
    firstName=fn.get()
    middleName=mn.get()
    lastName=ln.get()
    fullname = lastName+firstName+middleName
    fullname = fullname.upper()
    fullname = fullname.replace(" ", "")
    #Opening Excel
    theFile = openpyxl.load_workbook('Data/Search.xlsx')
    theFile1 = theFile.active
    for col in theFile1.iter_cols(1,5): 
        for row in range(1,25):
            if col[row].value == fullname:
                results = "This person is in group " + str(col[0].value)
                tkinter.messagebox.showinfo("Results",results)
                return
    if fullname == "" or cn(fullname) :
            tkinter.messagebox.showerror("Empty","Enter a valid name")
            return
    tkinter.messagebox.showerror("Not found","Student not found")


# Creating a Labels and Entries and button
fnl = tk.Label(root, text = 'First name (Imię) :', font=('calibre',10, 'bold'),bg='white')
fne = tk.Entry(root,textvariable = fn, font=('calibre',10,'normal'),bg='#EDEEF0')
mnl = tk.Label(root, text = 'Middle name (Drugie imię) :', font=('calibre',10, 'bold'),bg='white')
mne = tk.Entry(root,textvariable = mn, font=('calibre',10,'normal'),bg='#EDEEF0')
lnl = tk.Label(root, text = 'Last name (Nazwisko) :', font = ('calibre',10,'bold'),bg='white')
lne=tk.Entry(root, textvariable = ln, font = ('calibre',10,'normal'),bg='#EDEEF0')
sbtn=tk.Button(root, activebackground="#1750AC" ,bg="#003396", command=search ,width=15, height=1,highlightcolor="#86CEFA",text="Search",font=('calibre',10,'bold'),fg="white")


# GUI grid
fnl.grid(row=0,column=0,padx=(10,10),pady=(20,10))
fne.grid(row=0,column=1,pady=(20,10))
mnl.grid(row=1,column=0,padx=(10,10),pady=(10,10))
mne.grid(row=1,column=1,pady=(10,10))
lnl.grid(row=2,column=0,padx=(10,10),pady=(10,10))
lne.grid(row=2,column=1,pady=(10,10))
sbtn.grid(column=0,row=3,columnspan=3,padx=(28,0),pady=(15,0))



root.mainloop()