import folium
import os
cwd = os.getcwd()
html_file = os.path.join(cwd,"map.html")
# Make an empty map
m = folium.Map(location=[19.04,105.26],zoom_start=7.5)
# Import the pandas library
from folium.features import DivIcon

def number_DivIcon(color,number):
    """ Create a 'numbered' icon
    
    """
    icon = DivIcon(
            icon_size=(100,30),
            icon_anchor=(14,40),
            #html='<div style="font-size: 18pt; align:center, color : black">' + '{:02d}'.format(num+1) + '</div>',
            html="""<span class="fa-stack " style="font-size: 12pt">
                    <!-- The icon that will wrap the number -->
                    <span class="fa fa-circle-o fa-stack-2x" style="color : {:s}"></span>
                    <!-- a strong element with the custom content, in this case a number -->
                    <strong class="fa-stack-1x">
                         {:.1f}  
                    </strong>
                </span>""".format(color,number)
        )
    return icon
# Make a data frame with dots to show on the map
import pandas as pd
data = pd.read_csv("./Get_data_from_web/loc_name.csv")
div = folium.DivIcon(html=(
    '<svg height="100" width="100">'
    '<circle cx="50" cy="50" r="40" stroke="yellow" stroke-width="3" fill="none" />'
    '<text x="30" y="50" fill="black">9000</text>'
    '</svg>'
))
for i in range(0,len(data)-200):
   folium.Marker(
      location=[data.iloc[i]['lat'], data.iloc[i]['lon']],
      popup=data.iloc[i]['Station'],icon=folium.Icon(color='blue',icon_color='blue'),markerColor='blue',
   ).add_to(m)
   folium.Marker(
        location=[data.iloc[i]['lat'], data.iloc[i]['lon']],
        popup="Delivery " + '{:.2f}'.format(data.iloc[i]['lat']),
        icon= number_DivIcon('black',data.iloc[i]['lat'])
    ).add_to(m)
m.save(html_file)
import webview
# define an instance of tkinter
#  size of the window where we show our website
# Open website
webview.create_window('Bando',url=html_file,width=1024,height=768)
webview.start()