import requests, json
import xlwings as xw
import time


def get_temperature(city_name):
    url = "http://api.openweathermap.org/data/2.5/weather?appid=d53280578bb32d47d66ab966197e5574&q="
    complete_url = url+city_name
    req1 = requests.get(complete_url)
    data = req1.json()
    return data['main']['temp']


def kelvin_to_celsius(tmp):
    c =  tmp - 273.15
    return int(c)

def kelvin_to_farenheit(tmp):
    f = ((tmp - 273.15) * 1.8) + 32
    return int(f)



def weather_xlsx():
    workbook = xw.Book('/Users/ASUS/Desktop/LiveWeather/weather.xlsx')
    cities = workbook.sheets['Sheet1']
    
    values = workbook.sheets['Sheet2']
    
    while True:
        
        if cities['B4'].value == 1:
            current_temp = get_temperature("Kolkata")
            cel = kelvin_to_celsius(current_temp)
            far = kelvin_to_farenheit(current_temp)
            if cities['B3'].value == 'C':
                cities['B2'].value = cel
            elif cities['B3'].value == 'F':
                cities['B2'].value = far
            values['A'+str(values.range('A' + str(values.cells.last_cell.row)).end('up').row + 1)].value = str(cel)+' C/'+str(far)+' F'
            

        if cities['C4'].value == 1:
            current_temp = get_temperature("Delhi")
            cel = kelvin_to_celsius(current_temp)
            far = kelvin_to_farenheit(current_temp)
            if cities['C3'].value == 'C':
                cities['C2'].value = cel
            elif cities['C3'].value == 'F':
                cities['C2'].value = far
            values['B'+str(values.range('B' + str(values.cells.last_cell.row)).end('up').row + 1)].value = str(cel)+' C/'+str(far)+' F'
            
        if cities['D4'].value == 1:
            current_temp = get_temperature("Mumbai")
            cel = kelvin_to_celsius(current_temp)
            far = kelvin_to_farenheit(current_temp)
            if cities['D3'].value == 'C':
                cities['D2'].value = cel
            elif cities['D3'].value == 'F':
                cities['D2'].value = far
            values['C'+str(values.range('C' + str(values.cells.last_cell.row)).end('up').row + 1)].value = str(cel)+' C/'+str(far)+' F'
            
        time.sleep(2) #updates value every 2 seconds


weather_xlsx()