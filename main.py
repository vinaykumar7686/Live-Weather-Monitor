# coding: utf-8
import requests
from pathlib import Path
import openpyxl, os,sys, time

def toFah(cel):
  # print(f'B:{cel}')
  return (cel*9//5)+32

def get_Value(loc):
  __params = {
  'access_key': 'efa59c8305d828986c91c19480a592a5',
  'query': loc
  }
  api_result = requests.get('http://api.weatherstack.com/current', __params)

  api_response = api_result.json()

  # print(api_response)

  return [api_response['request']['query'], api_response['current']['humidity'], api_response['current']['temperature']]


def get_ques():
  xlsx_file = Path('D:\Programming\Python\Projects\Live Weather Monitor\Values.xlsx')
  wb_obj = openpyxl.load_workbook(xlsx_file) 
  sheet = wb_obj.active
  initial = True
  n = sheet.max_row
  while True:
    for i in range(1, n):

      if (sheet.cell(row = i+1, column = 5).value) == 1 or initial:

        city = sheet.cell(row = i+1, column = 1).value
        values = get_Value(city)

        if (sheet.cell(row = i+1, column = 4).value) == 'C':
          sheet.cell(row = i+1, column = 2).value =  values[2]
          sheet.cell(row = i+1, column = 3).value =  values[1]
          wb_obj.save("D:\Programming\Python\Projects\Live Weather Monitor\Values.xlsx") 
          print(f'Loc: {values[0]}')
          print(f'Humidity: {values[1]}')
          print(f'Temp: {values[2]} C')
        
        else:
          sheet.cell(row = i+1, column = 2).value =  toFah(values[2])
          sheet.cell(row = i+1, column = 3).value =  values[1]
          wb_obj.save("D:\Programming\Python\Projects\Live Weather Monitor\Values.xlsx") 
          print(f'Loc: {values[0]}')
          print(f'Humidity: {values[1]}')
          print(f'Temp: {toFah(values[2])} F')

      time.sleep(1)
    initial = False


if __name__ == "__main__":
    get_ques()
