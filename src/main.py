import xlsxwriter

def menu():
    #Menu 
    print("-----------------------------ERPMate--------------------------------")
    print("1. Start an ERP session")
    print("2. Exit")
    menyo=int(input("Enter your choice: "))
    if menyo==1:
        ERPsession()
    if menyo==2:
        pass

def ERPsession():
    #start of an erp session
    print("------------------------------ERP Session----------------------------")
    #making variables global
    global date,time,obssession,compulsion,anxietyatexposure,anxietyat5min,anxietyat10min,anxietyat30min,anxietyat60min
    #asking for input
    date=input("Enter today's date: ")
    time=input("Enter current time: ")
    obssession=input("Enter the obsessive thought: ")
    compulsion=input("Enter the compulsion: ")
    print("------------------Rate your anxiety at 10---------------")
    anxietyatexposure=input(" Rate your Anxiety at Exposure: ")
    anxietyat5min= input("Rate your Anxiety after 5 min: ")
    anxietyat10min= input("Rate your Anxiety after 10 min: ")
    anxietyat30min= input("Rate your Anxiety after 30 min: ")
    anxietyat60min= input("Rate your Anxiety after 1 hour: ")
    createExcelsheet()
    print("Data recorded in excel sheet")

def createExcelsheet():
    #creating an excel sheet
    f=input("Enter the file name in the format 'filename.xlsx': ")
    workbook=xlsxwriter.Workbook(f)
    worksheet=workbook.add_worksheet("ERPData")
    #adding columns
    worksheet.write("A1","Date")
    worksheet.write("B1", "Time")
    worksheet.write("C1","Obsession")
    worksheet.write("D1","Compulsion")
    worksheet.write("E1","Anxiety at Exposure")
    worksheet.write("F1","Anxiety after 5 min")
    worksheet.write("G1","Anxiety after 10 min")
    worksheet.write("H1","Anxiety after 30 min")
    worksheet.write("I1","Anxiety after 1 hour")

    #small workaround
    session = int(input("Enter the session number: "))
    session = session + 1
    sessionzero=str(session)
    
    #adding the session data to excel sheet
    worksheet.write("A"+sessionzero,date)
    worksheet.write("B"+sessionzero,time)
    worksheet.write("C"+sessionzero,obssession)
    worksheet.write("D"+sessionzero,compulsion)
    worksheet.write("E"+sessionzero,anxietyatexposure)
    worksheet.write("F"+sessionzero,anxietyat5min)
    worksheet.write("G"+sessionzero,anxietyat10min)
    worksheet.write("H"+sessionzero,anxietyat30min)
    worksheet.write("I"+sessionzero,anxietyat60min)
             
    workbook.close()

    

    






menyoo=1
#While loop for menu
while menyoo==1:
    menu()
