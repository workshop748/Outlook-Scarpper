import pandas as pd
import datetime as dt
import win32com.client as wc
import pythoncom

#a function to get the item property
def getItemProperty(item, propertyName):
    try:
        result =getattr(item,propertyName)
    except pythoncom.com_error as ce:
        result =ce.excepinfo[2]
    return result    

#formate to getjust the date
def formatDate(theday):
    seased=str(theday)
    theday =seased.split(' ')
    nosences=theday[0].split("-")
    print(nosences[2])
    return nosences[2]
#setting up the outlook calander data
def getOutLook_calander():
   #this is going to open up the current outlook application 
   Outlook = wc.Dispatch("Outlook.Application")
   print("Outlook version: {}".format(Outlook.Version))
   print("Default profile name: {}".format(Outlook.DefaultProfileName))
    #getting the namespace object
   namespace =Outlook.Session
   #this will be edited when I have access to the resnet email.
   recipient =namespace.createRecipient("resnet@uww.edu")
   sharedCalender = namespace.GetSharedDefaultFolder(recipient,9)
   return sharedCalender

def get_collectedData(start, end, clander):
    #this section is used to take the stored data and store it in a pandas data structure
    item = clander.Items
    restriction = restriction = "[Start] >= '" + start.strftime("%m/%d/%Y") + "' AND [End] <= '" +end.strftime("%m/%d/%Y") + "'"
    restrictionItems= item.Restrict(restriction)
    columns = ['Start Time', 'End Time','Location','Subject','Day']
    df =pd.DataFrame(columns=columns)
    for things in restrictionItems:
        #this will itereate through items and will store in a pandas dictinary
        startDate = getItemProperty(things, "Start")
        endDate = getItemProperty(things, "End")
        location= getItemProperty(things,"Location")
        startTime= things.Start.strftime('%H:%M')
        endTime = things.End.strftime('%H:%M')
        subject = getItemProperty(things, "Subject")

        startDate =formatDate(startDate)
        endDate =formatDate(endDate)
        day =startDate
        #storing the items into a pandas dictinary.
        df = df._append([{
            'Start Time': startTime,
            'End Time': endTime,
            'Location': location,
            'Subject': subject,
            'Day':day
                       }],ignore_index=True)
        
    
    
    return df

DaySeased= dt.date.today()
WeekDay = DaySeased.isoweekday()
#the start of the week day
start = DaySeased -dt.timedelta(days=WeekDay)
# build a simple range
dates =[start+dt.timedelta(days=d) for d in range(7)]

for days in dates:
    aDay =str(days)
    storedDates=aDay.split("-")
    year, month, day = int(storedDates[0]),int(storedDates[1]),int(storedDates[2])
    if days == DaySeased:
        begin=dt.datetime(year,month,day)
    elif days == dates[-1]:
       end =dt.datetime(year,month,day)

sharedCalender=getOutLook_calander()
TheStolenData = get_collectedData(start,end,sharedCalender)
print(TheStolenData)
TheStolenData.to_csv('appointment_hours.csv')