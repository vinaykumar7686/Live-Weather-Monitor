# coding: utf-8
import requests

params = {
  'access_key': 'efa59c8305d828986c91c19480a592a5',
  'query': 'New York'
}

api_result = requests.get('http://api.weatherstack.com/current', params)

api_response = api_result.json()

#print(api_response)

print(u'Current temperature in %s is %d℃' % (api_response['location']['name'], api_response['current']['temperature']))