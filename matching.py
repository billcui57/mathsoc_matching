import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint

scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name(
    "MathSoc Job Matching-0e05c076907d.json", scope)

client = gspread.authorize(creds)

spreadSheet = client.open("Matching Data")

worksheets = spreadSheet.worksheets()




print("Clearing previous results...\n")


for worksheet in worksheets[:-1]:
    cell_list = worksheet.range("D2:G100")
    for cell in cell_list:
        cell.value = ''
    worksheet.update_cells(cell_list)





interviewers = []

print("Reading raw data...\n")

#Reads in all the interviewers

for x in worksheets[:-1]:
    interviewers.append(
        {
            "name": x.cell(2, 1).value,
            "program": x.cell(2, 2).value,
            "pastjobs": x.col_values(3)[1:],
            "intervieweeTimes": [],
            "interviewees": [],
            "intervieweesStudNumber": [],
            "intervieweesEmail": [],
            "researchAssistantRate": 0,
            "teacherRate": 0,
            "financialAnalystRate": 0,
            "statisticianRate": 0,
            "softDeveloperRate": 0,
            "machineLearningInternRate": 0,
            "dataScientistRate": 0,
            "accountantRate": 0,
            "businessRepRate": 0,
            "mathModelerRate": 0,
            "marketAssistantRate": 0,
            "managerRate": 0
        }
    )



print("Assigning expertise values to each interviewer...\n")

# Gives rating to the level of expertise for each interviewer in each skill
# TODO: Let major affect level too

for interviewer in interviewers:
    for job in interviewer["pastjobs"]:
        if job.lower() == "Research Assistant".lower():
            interviewer["researchAssistantRate"] += 1
        elif job.lower() == "Teacher".lower():
            interviewer["teacherRate"] += 1
        elif job.lower() == "Financial Analyst".lower():
            interviewer["financialAnalystRate"] += 1
        elif job.lower() == "Statistician".lower():
            interviewer["statisticianRate"] += 1
        elif job.lower() == "Software Developer".lower():
            interviewer["softDeveloperRate"] += 1
        elif job.lower() == "Machine Learning Intern".lower():
            interviewer["machineLearningInternRate"] += 1
        elif job.lower() == "Data Scientist".lower():
            interviewer["dataScientistRate"] += 1
        elif job.lower() == "Accountant".lower():
            interviewer["accountantRate"] += 1
        elif job.lower() == "Business Representative".lower():
            interviewer["businessRepRate"] += 1
        elif job.lower() == "Mathematical Modeler".lower():
            interviewer["mathModelerRate"] += 1
        elif job.lower() == "Marketing Assistant".lower():
            interviewer["marketAssistantRate"] += 1
        elif job.lower() == "Manager".lower():
            interviewer["managerRate"] += 1
        else:
            print("Error parsing", job)
     


print("Reading interviewees...\n")

#Reads in interviewees

interviewees = []

formSheet = worksheets[len(worksheets)-1]

for row in range(1,formSheet.row_count):

    if not formSheet.cell(row+1, 3).value:
        break
    
    interviewees.append(
        {
            "name": formSheet.cell(row + 1, 3).value + " " + formSheet.cell(row + 1, 4).value,
            "email": formSheet.cell(row + 1, 2).value,
            "studentNum": int(formSheet.cell(row + 1, 5).value),
            "program ": formSheet.cell(row + 1, 6).value,
            "time": formSheet.cell(row + 1, 20).value,
            int (formSheet.cell(row + 1, 7).value) : "researchAssistantRate",
            int (formSheet.cell(row + 1, 8).value) : "teacherRate",
            int (formSheet.cell(row + 1, 9).value) : "financialAnalystRate",
            int (formSheet.cell(row + 1, 10).value): "statisticianRate",
            int (formSheet.cell(row + 1, 11).value): "softDeveloperRate",
            int (formSheet.cell(row + 1, 12).value): "machineLearningInternRate",
            int (formSheet.cell(row + 1, 13).value): "dataScientistRate",
            int (formSheet.cell(row + 1, 14).value): "accountantRate",
            int (formSheet.cell(row + 1, 15).value): "businessRepRate",
            int (formSheet.cell(row + 1, 16).value): "mathModelerRate",
            int (formSheet.cell(row + 1, 17).value): "marketAssistantRate",
            int (formSheet.cell(row + 1, 18).value): "managerRate",
        }
    )
    



print("Matching...\n")

#Max capacity of each interviewer
#TODO make it more robust
maxCapacity = int(len(interviewees) / len(interviewers) + 5) 

qualification_threshold = 0

#Algorithm for matching
for interviewee in interviewees:
    desiredJobRank = 1
    matchedInterviewer = False
    while True:
        for interviewer in interviewers:
            if (interviewer[interviewee[desiredJobRank]] > qualification_threshold) and (len(interviewer["interviewees"])) < maxCapacity:
                interviewer["interviewees"].append(interviewee["name"])
                interviewer["intervieweesStudNumber"].append(interviewee["studentNum"])
                interviewer["intervieweesEmail"].append(interviewee["email"])
                interviewer["intervieweeTimes"].append(interviewee["time"])
                matchedInterviewer = True
                break
            else:
                continue
               
        
        if not matchedInterviewer:
            desiredJobRank += 1
        else:   
            break
             


#Pushes changes back (overrides and does not appends)

print("Pushing match results...\n")


for worksheet in worksheets[:-1]:
    interviewerIndex = worksheets.index(worksheet)
    for row in range(1,len(interviewers[interviewerIndex]["interviewees"])+1):
        worksheet.update_cell(row+1,4,interviewers[interviewerIndex]["interviewees"][row-1])
        worksheet.update_cell(row+1,5,interviewers[interviewerIndex]["intervieweesStudNumber"][row-1])
        worksheet.update_cell(row+1,6,interviewers[interviewerIndex]["intervieweesEmail"][row-1])
        worksheet.update_cell(row+1,7,interviewers[interviewerIndex]["intervieweeTimes"][row-1])
    


print("Done!\n")

