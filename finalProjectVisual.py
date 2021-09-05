import openpyxl
from tkinter import *
wb = openpyxl.load_workbook("netflix_titles.xlsx")
sheet = wb["netflixtitles"]
#CREATES MAIN SCREEN
root = Tk()
root.geometry("400x400")
root.title("Movie Selection")
root["bg"]="green"
#CREATES FRAME IN MAIN SCREEN
frame = Frame(root)
#CREATES lABEL
lbl1 = Label(root,text = "Choose A Movie",borderwidth=20,bg="green",fg="white",font=15)
lbl1.pack()
#CREATES SCROLLBAR
myScrollBar = Scrollbar(frame, orient=VERTICAL)
myListBox = Listbox(frame,width = 50,yscrollcommand=myScrollBar.set)
myScrollBar.config(command=myListBox.yview)
myScrollBar.pack(side=RIGHT,fill=Y)
frame.pack()
myListBox.pack()
#---------------------CLASS-------------------
class Movie:
    def __init__(self,type,title,director,releaseYear):
        self.type = type
        self.title = title
        self.director = director
        self.releaseYear = releaseYear
    
def createMovie(row):#Creates Movie and moves it to Class
    mov = Movie(row[0].value,row[1].value,row[2].value,row[3].value)
    return mov

def get_data_from_excel():#Pulls Data from Excel
    movies = []
    for row in sheet["A2:d25"]:
        movies.append(createMovie(row))
    return movies

def showTitle():#Creates List Box
    myListBox.bind("<ButtonRelease-1>",select)#<ButtonRelease-1> creates a single click on mouse
    for title in movies:
        myListBox.insert(END,title.title)
def select(e):#ACCEPTS BIND
    for mov in movies:
        dict = {mov.title:"Director: "+ mov.director +" "+ " \nRelease Year: "+ str(mov.releaseYear)}
        lbl2.config(text=dict.get(myListBox.get(ANCHOR)))

    
#CALLS FUNCTION TO ACTION
movies = get_data_from_excel()
showTitle()

#CREATES EMPTY LABEL SO WE CAN REPALCE IT
lbl2 =Label(root,text="",borderwidth=15,bg="green",fg="white",font=15)
lbl2.pack()





root.mainloop()