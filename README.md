# RoadMapMaker
This gs enable you to create an interactive roadmap that can be used for keep track of your projects.
The RMM_Form script uses a form to trigger the script. 


1. Copy and paste the complete code to a clean appScriptfile.gs
2. Create a form, Example form -> https://docs.google.com/forms/d/e/1FAIpQLSc8bXCy6mjV8Q2RKyhHlgnYIDTNtFBDfwmKzsMcYb4VnjzdTw/viewform
3. the form needs 2 fields Date field and Choice Field (with number 1,2,3,6,12) month.
4. The google apps script RMM_SCRIPT is triggered once the form is submitted. 
   Therefore Ensure to setup a trigger to execute as soon as the form is submitted. Configure 'Trigger' section to run  this FormSubmit() function
5. Note The google sheets seems to be a special type of 'form' i.e a sheet that's linked to a form
    example results spreadsheet -> https://docs.google.com/spreadsheets/d/1dnDKfkxq5-CyPILTOx5gn4GngeCQj6MNkyH98SpdVtA/edit?resourcekey#gid=1807623523
6. The google Apps Script code is -->> https://script.google.com/home/projects/1Z9_mqrp3Gow-P3Q4xqvRsezbneGSjIQrm2btudPmv531K5C_qU3BgmTs/edit
