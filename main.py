# coding: utf-8
import requests

def get_Value(loc):
  __params = {
  'access_key': 'efa59c8305d828986c91c19480a592a5',
  'query': loc
  }
  api_result = requests.get('http://api.weatherstack.com/current', __params)

  api_response = api_result.json()

  # print(api_response)
  # print(f"Humidity:{api_response['current']['temperature']} Temperature:{api_response['current']['humidity']}")

  return [api_response['request']['query'], api_response['current']['humidity'], api_response['current']['temperature']]

print(get_Value('India'))


