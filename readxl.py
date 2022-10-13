import string
import xlwings as xw
import time
from os import system
import matplotlib.pyplot as plt

#Locate the xlsx file and specific sheet (Make sure editing is enabled for the file)
ws = xw.Book('PresResults.xlsx').sheets["2016 General"]

#Read the values from the whole sheet into a list of lists
table = ws.range("A1:F70").value


#Present a menu to choose option 1 or 2 or 3
#Option 1: Display a list of counties with an associated number id
#Option 2: Select a county to view results
#Option 3: Exit

#Create a list to associate each county with an id
countynames = [
'ADAMS', 'ALLEGHENY',
 'ARMSTRONG', 'BEAVER', 'BEDFORD', 'BERKS', 'BLAIR', 'BRADFORD',
 'BUCKS', 'BUTLER', 'CAMBRIA', 'CAMERON', 'CARBON', 'CENTRE',
 'CHESTER', 'CLARION', 'CLEARFIELD', 'CLINTON', 'COLUMBIA',
 'CRAWFORD', 'CUMBERLAND', 'DAUPHIN', 'DELAWARE', 'ELK', 'ERIE',
 'FAYETTE', 'FOREST', 'FRANKLIN', 'FULTON', 'GREENE', 'HUNTINGDON',
 'INDIANA', 'JEFFERSON', 'JUNIATA', 'LACKAWANNA', 'LANCASTER', 
 'LAWRENCE', 'LEBANON', 'LEHIGH', 'LUZERNE', 'LYCOMING', 'McKEAN',
 'MERCER', 'MIFFLIN', 'MONROE', 'MONTGOMERY', 'MONTOUR', 'NORTHAMPTON', 
 'NORTHUMBERLAND', 'PERRY', 'PHILADELPHIA', 'PIKE', 'POTTER', 'SCHUYLKILL', 'SNYDER',
 'SOMERSET', 'SULLIVAN', 'SUSQUEHANNA', 'TIOGA', 'UNION', 'VENANGO', 'WARREN', 'WASHINGTON',
 'WAYNE', 'WESTMORELAND', 'WYOMING', 'YORK'
 ]



#Create the menu function which runs on startup
def Menu():
    system('cls')
    print("PA Presidetial Election by County\n")
    print("Option 1: View list of county names")
    print("Option 2: Select county to view results")
    print("Option 3: Exit")

#Prompt user for input
    selection = input("Enter your choice: ")

    if selection.isalpha():
        Menu()
    if int(selection) == 1: #Display county list
        listCounties()

    elif int(selection) == 2: #Select county to view results
        viewCountyResults()
    elif int(selection) == 3: #Exit the app
        print("Thank you for using our App!")
        time.sleep(2)
        quit()
    elif int(selection) != 1 or 2 or 3:
        Menu()

#Create function to list out the counties
def listCounties():
    system('cls')
    for id, name in enumerate(countynames): #Loop through the list and print a numerical ID by their name
        print(id, name)
    #Ask user if they want to return    
    retry = input("Return to menu? (y/n): ") 
    if retry == 'y':
         Menu()
    elif retry == 'n':
        print("Thanks for enjoying our App!")
        time.sleep(3)
        system('cls')
        quit()
    elif retry != 'y' or 'n':
        Menu()

        
#Create a function to view the county results
def viewCountyResults():
    #NEEDS VALIDATION
    countyChoice = int(input("Input a county ID: ")) #Select a county by its ID 
    countyName = countynames[countyChoice] #Grab the name of the county by its ID
    get_info(table,countyName)#Pass in Excel table data & name of county 
    


#Create a function to loop through the data & create the bar graph
def get_info(table, query):
    for row in table:#Loop through each list
        if row[0] == query:#Check if the first value of each list equals the county name
            results = [row[1], row[2]] #Store the votes for hillary & trump
            candidates = ["Hillary Clinton", "Donald Trump"] #Store Hillary & Trumps name
            plt.bar(candidates, results)#Create a bar graph using the candidate names & vote data
            plt.title(query) #Title of graph is the county name
            plt.xlabel("Candidates")
            plt.ylabel("# of votes")
            plt.show()#Display the graph
            Menu()#Restart the app on graph close
            
Menu()