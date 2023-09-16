import folium
import sys,os
cwd = os.getcwd()
html_file = os.path.join(cwd,"map.html")
# Make an empty map
m = folium.Map(location=[18.1,105.3],zoom_start=7.5)

# Import the pandas library
import pandas as pd

# Make a data frame with dots to show on the map
data = pd.read_csv("./Get_data_from_web/loc_name.csv")

for i in range(0,len(data)):
   folium.Marker(
      location=[data.iloc[i]['lat'], data.iloc[i]['lon']],
      popup=data.iloc[i]['Station'],
   ).add_to(m)
   m.save(html_file)
import webview
# define an instance of tkinter
#  size of the window where we show our website
# Open website
webview.create_window('Bando',url=html_file,width=1024,height=768)
webview.start()