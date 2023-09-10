def dir_conver(str):
   switcher = {
      '00': '','02': 'NNE',5: 'NE',7: 'ENE', 9: 'E',
      11: 'ESE', 14: 'SE', 16: 'SSE', 18: 'S',
      20: 'SSW', 23: 'SW', 25: 'WSW', 27: 'W',
      29: 'WNW', 32: 'NW', 34: 'NNW', 36: 'N'
      }
   return switcher.get(str,"invalid wind direction")
print(dir_conver())
