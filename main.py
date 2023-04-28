import requests
import os
from docx import Document
from docx.shared import Inches
from docx.shared import RGBColor
import docx

VIRTUOSO_API = "api-app2.virtuoso.qa"
VIRTUOSO_API_EXECUTION_DETAILS_PATH = "executions"
SCREENSHOT_FOLDER = "screenshots"
# Change this according to your needs
VIRTUOSO_TOKEN = "d5b93783-7b49-4e13-941c-4dafe30c2cf9"
VIRTUOSO_EXECUTION_ID = 30125


# Retrieve the screenshots from the execution details
executionRequestURL = "https://{}/api/testsuites/execution?jobId={}".format(VIRTUOSO_API,
                                                    VIRTUOSO_EXECUTION_ID)

executionRequestHeaders = {"Authorization": "Bearer {}".format(VIRTUOSO_TOKEN)}

executionDetails = requests.get(executionRequestURL, headers=executionRequestHeaders)

if executionDetails.status_code != 200:
    print("Unexpected response code: {}".format(executionDetails.status_code))
    exit(1)

# Check if screenshots folder exists
if not os.path.exists(SCREENSHOT_FOLDER):
    os.makedirs(SCREENSHOT_FOLDER)

#document = Document()
doc = docx.Document()
doc.add_paragraph().add_run('Steps Screenshots').bold = True
doc.add_paragraph().add_run('\n')


goals = executionDetails.json().get('item', {}).get('journeys', {})

allsteps = []
screenshots = []
i = 0
for goal in goals:

    journeydetail = goals[goal].get('journey', {})
    snapshotId = journeydetail['snapshotId']
    goalId = journeydetail['goalId']

    stepsURL = "https://{}/api/snapshots/{}/goals/{}/testsuites".format(VIRTUOSO_API,snapshotId, goalId)
    stepsHeaders = {"Authorization": "Bearer {}".format(VIRTUOSO_TOKEN)}
    stepsDetails = requests.get(stepsURL, headers=stepsHeaders)
  
    maps = stepsDetails.json().get('map', {})
    steps = []
    for mapdetail in maps:
        cases = maps[mapdetail].get('cases', {})
        for case in cases:
            for step in case['steps']:
                allsteps.append(step)

    stepstext = {}
    stepstextURL = "https://step-deparser-service.virtuoso.workers.dev"
    stepsHeaders = {"Authorization": "Bearer {}".format(VIRTUOSO_TOKEN)}
    stepstextDetails = requests.post(stepstextURL, json={'steps': allsteps})

    j = 0
    for allstep in allsteps:
        stepstext[allstep['id']] = stepstextDetails.json()[j]
        j = j+1
        

    journeys = goals[goal].get('lastExecution', {}).get('report', {}).get('checkpoints', {})
    for journey in journeys:
        steps = journeys[journey].get('steps', {})

        
        for step in steps:
            i = i+1
            #if(i<=8):
            
            stepId = steps[step].get('stepId')
            doc.add_paragraph().add_run('Step : '+stepstext[stepId])

            beforescreenshot = steps[step].get('beforeScreenshot')
            if(beforescreenshot):

                # Save screenshots to fs
                beforescreenshotFileName = beforescreenshot.split('/')[-1]
                beforescreenshotFile = requests.get(beforescreenshot)

                open('{}/{}'.format(SCREENSHOT_FOLDER, beforescreenshotFileName), 'wb').write(beforescreenshotFile.content)
                
                doc.add_paragraph().add_run('Before Execution Screenshot')
                doc.add_picture(SCREENSHOT_FOLDER+'/'+beforescreenshotFileName,width=Inches(6.0), height=Inches(7.8))
                doc.add_paragraph().add_run('\n')

                print('saving image')


            screenshot = steps[step].get('screenshot')
            if(screenshot):

                # Save screenshots to fs
                screenshotFileName = screenshot.split('/')[-1]
                screenshotFile = requests.get(screenshot)
                open('{}/{}'.format(SCREENSHOT_FOLDER, screenshotFileName), 'wb').write(screenshotFile.content)
                
                doc.add_paragraph().add_run('After Execution Screenshot')
                if(i == 1): 
                    doc.add_picture(SCREENSHOT_FOLDER+'/'+screenshotFileName,width=Inches(6.0), height=Inches(7.2))
                else :
                    doc.add_picture(SCREENSHOT_FOLDER+'/'+screenshotFileName,width=Inches(6.0), height=Inches(8))

                print('saving image')

                
            



doc.save('Steps Screenshots.docx')
