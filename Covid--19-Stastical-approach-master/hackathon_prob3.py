#SHRINIKETAN S TIKARE(19BEC041)
# BEFORE RUNNING THIS PROGRAM PLEASE INSTALL PYTHON MODULE CALLED pywin32,u can do this by typing-<pip install pywin32> command in command prompt(for windows) or bash (mac terminal) 
from win32com.client import Dispatch
import pandas as pd
import csv

def speak(str):
    speak=Dispatch("SAPI.SPVoice")
    speak.Speak(str)
    
class change_country_state(object):
    def __init__(self,country_name,state):
        self.country_name=country_name
        self.state=state
    def finding_avg_no_of_days_for_country_to_change(self):
        state1 = pd.read_csv(f"G:\\phytho\\uma mam\\hackathon\\group hackathon\\Covid--19-Stastical-approach-master\covid19_{self.state}_global.csv")#recovered
        state1 = state1.drop(['Province/State','Lat','Long'],axis=1)
        state1 = state1.groupby(state1['Country/Region']).aggregate('sum')
        state1=state1.T
        no_of_days=0
        for i in state1[self.country_name]:
            if i==0:
                no_of_days+=1
        print(f"It took {no_of_days} days to change the state of country to {self.state} from the date 22nd january 2020.")
        speak(f"It took {no_of_days} change the state of country to {self.state} from the date 22nd january 2020.")
  
def max_no_of_days_it_took_by_country_in_whole_world():
        death1 = pd.read_csv(f"G:\\phytho\\uma mam\\hackathon\\group hackathon\\Covid--19-Stastical-approach-master\covid19_deaths_global.csv")#recovered
        death1 = death1.drop(['Province/State','Lat','Long'],axis=1)
        death1 = death1.groupby(death1['Country/Region']).aggregate('sum')
        death1=death1.T

        lst={}
        for item1 in death1.columns:
            no_of_days=0
            for item2 in death1[item1]:
                if item2==0:
                    no_of_days+=1
            lst[no_of_days]=item1
            
        lst1=[]
        for item3 in lst:
            lst1.append(item3)
        
        print(f"{lst[max(lst1)]} is the country, that has taken maximum number of days to come to death state, and the days taken by it are {max(lst1)} number of days")
        speak(f"{lst[max(lst1)]} is the country, that has taken maximum number of days to come to death state, and the days taken by it are {max(lst1)} number of days")
        
        

if __name__ == '__main__':
    
    speak("Hello everyone, this is COVID-19 visualisar. Please enter the country that you want the information")
    loop=1
    while(loop==1):
        country_name=input("Please enter the country that you want the information:").capitalize()

        contiinue=1
        while(contiinue==1):
            speak("enter the state you need to find out average days taken to a country ")
            state=int(input("enter the state you need to find out average days taken to a country to turn in\n1.death\n2.recovered\n"))

            if state==1:
                   death=change_country_state(country_name,"deaths")
                   death.finding_avg_no_of_days_for_country_to_change()
            if state==2:
                   recovered=change_country_state(country_name,"recovered")
                   recovered.finding_avg_no_of_days_for_country_to_change()
            speak("do u want to continue looping into same country? Press 1 to do so,, else press any key to exit out ")     
            contiinue=int(input("do u want to continue looping into same country? Press 1 to do so,, else press any key to exit out "))
        
        speak("do u want to explore the data of other country? Press 1 to do so,, else press any key to exit out ")
        loop=int(input("do u want to explore the data of other country? Press 1 to do so,, else press any key to exit out "))

    max_no_of_days_it_took_by_country_in_whole_world()
    