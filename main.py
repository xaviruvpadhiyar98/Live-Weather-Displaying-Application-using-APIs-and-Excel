# imports

import pandas as pd
from time import sleep
from requests import get
from requests.utils import quote

# Unit Conversion Func
def checkunit(x,t):
    if 'c' == x.lower():
        return int(abs(t - celsius)),'C'
    if 'f' == x.lower():
        return int(abs(t - fahrenheit)),'F'
    if 'k' == x.lower():
        return int(t), 'K'


# PreDefined Variables

filename = 'Live-Weather.xlsx'
url = 'https://api.openweathermap.org/data/2.5/weather?appid=c12cca9e48ca1f6a195c9178adf4349b&q='
fahrenheit = 459.67
celsius = 273.15


# Run Infinite Times
while True:
	
	# Loading the excel File
	df = pd.read_excel(filename, index_col=False)
	
	print(f'{filename} is been updating please wait....')
	# Processing the Excel File with pandas 
	results = []
	for index, row in df.iterrows():
		city = row['City']
		unit = row['Unit']
		update = row['Update']
		temperature = row['Temperature']
		humidity = row['Humidity']

		if update == 1:
			response = get(url+quote(city)).json()
			u =  checkunit(unit, response['main']['temp'])
			results.append({
				'City': response['name'],
				'Temperature': u[0],
				'Humidity': response['main']['humidity'],
				'Unit': u[1],
				'Update': update
			})
		else:
			results.append({
				'City': city,
				'Temperature': temperature,
				'Humidity': humidity,
				'Unit': unit,
				'Update': update
			})
			
		sleep(1)
		
	# Saving back to excel
	pd.DataFrame(results).to_excel(filename, index=False)
	
	print(f'{filename} is updated')
	print("\n Next Update will take place after 20 seconds")
	# Running this every 20 Seconds
	sleep(20)
