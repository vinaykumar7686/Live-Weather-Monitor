# Live-Weather-Monitor

This application is a live weather displaying application using APIs and excel to display and control the data being presented to the user.

- This application uses "http://api.weatherstack.com/current" a live weather API which provides us with live updating temperatures.

- We have used openpyxl to read and update the excel sheet cells at an interval of a few seconds.

- Description of excel file
    - City Name (use more than 1)
    - Temperature
    - Humidity
    - Option to change temperature (°C/°F)
    - Option to stop updating the temperature value(0/1)

- Generate your API access key by registering at  "http://api.weatherstack.com/current" and paste that key at line 41.<br>
__params = {
  'access_key': 'xxxxxxxxxxxxxxxxxxxxxxxxxxxx',
  'query': loc
  }
    
### How to Run?

- Install Clone the Repository
- Install requirements.txt by typing "pip install -r requirements.txt" into Command Prompt in working Directory.
- Edit the values of Cities, Units and Update(0/1) colulmns in the Values.xlsx file as required.
- Run the file main.py.
