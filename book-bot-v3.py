# -*- coding: utf-8 -*-
"""
Created on Mon Nov 29 21:57:46 2021

@author: Liya Thomas
"""

import PySimpleGUI as sg
import pandas as pd 
from datetime import date 

'enter in file name- ensure the file is already created'

fileName= "Book_Bot_2.0.xlsx"

#'define all layouts for GUI in functions'
#'layout to input new book values '
def inputBooks(): 

    layout = [[sg.Text('BOOK BOT 2.0')],
               [sg.Text('Title'), sg.InputText(key='t', do_not_clear=False)], 
               [sg.Text('Author'), sg.InputText(key='a', do_not_clear=False)],
                [sg.Text('Genre'), sg.InputText(key='g', do_not_clear=False)],
                [sg.Text('Comments'), sg.InputText(key='c', do_not_clear=False)],
                [sg.Button('Ok'), sg.Button('Exit'), sg.Button('Return')] ]
    return sg.Window('Book Bot 2.0', layout)


#'main GUI Layout'
def main():
    
    layout2= [[sg.Text('BOOK BOT 2.0')],
    
            [sg.Button('Enter New Book to Read'), sg.Button('View List'), sg.Button('Exit') ]]
    sg.theme('Dark Purple 1')
    return sg.Window('Book Bot 2.0', layout2)

##'book list to read layout '
def booksToread(x): 


      
    layout3 = [  [sg.Text('Books to Read List')],
                 [sg.Listbox(values=x, size=(30, len(x)+2), key='selection')],
                 [sg.Button('Exit')], 
                 [sg.Button('Return')]
                ]
   
     
    return (sg.Window('Book Bot 2.0', layout3) )

'


#'define all the lists used to document what values are being inputted in the GUI '
bTitle=[]
bAuthor=[]
bComment=[]
bGenre=[]


#'read the excel to populate the listbox for the book list to read layout'
data=pd.read_excel(fileName)
data_list= data['Title'].tolist() 

#define main window
window = main()

#contnious while true loop to always have an option to select 
while True:
#read each window and determine which event is being selected
    event, value = window.read()
    #exit if the window is closed or the value is exited 
    if event == sg.WIN_CLOSED or event == 'Exit': # if user closes window or clicks cancel
        break
    if event== 'Enter New Book to Read':
        window.close() 
        #redefine window value to make sure the event function still works
        #call function to enter new book
        window=inputBooks()
    if event== 'Ok': 
        #store the input values from the GUIs into the lists defeined earlier 
        bTitle.append(value['t'])
        bAuthor.append(value['a'])
        bComment.append(value['c'])
        bGenre.append(value['g'])
    if event=='View List':
        window.close()
        #call function to view list
        window=booksToread(data_list)
    if event== 'Return': 
        window.close() 
        #return to the main dashboard 
        window=main() 
    

'DATAFRAME BUILD' 
# assign inputs but in the user           
bookList={}
# assign the list values and add date value 
bookList = {'Date': date.today(), 'Title':bTitle, 'Author': bAuthor,'Genre':bGenre, 'Comment': bComment}
# define dataframe 
bDF=pd.DataFrame(bookList, columns=['Date', 'Title', 'Author','Genre', 'Comment'])

# read existing excel file to gather books already in there 
currentDF=pd.read_excel(fileName)

# concatnate the input dataframe dataframe and the existing excel file choices 
totalDF=pd.concat([bDF, currentDF])
# rename the completed dataframe 
finalDF=pd.DataFrame(totalDF, columns=['Date','Title', 'Author', 'Genre', 'Comment'])
# open the excel file to transfer values
writer = pd.ExcelWriter(fileName, engine='xlsxwriter')
# Convert the dataframe to an XlsxWriter Excel object.
finalDF.to_excel(writer, sheet_name='BTR')

writer.save()
            
window.close()
