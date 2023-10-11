import folium
import os
cwd = os.getcwd()
print (cwd)
html_file = os.path.join(cwd,"map.html")
# Make an empty map
m = folium.Map(location=[19.04,105.26],zoom_start=7.5)

# Import the pandas library
import pandas as pd
from folium.features import DivIcon

def number_DivIcon(color,number):
    """ Create a 'numbered' icon
    
    """
    icon = DivIcon(
            icon_size=(120,36),
            icon_anchor=(12,35),
            #html='<div style="font-size: 18pt; align:center, color : black">' + '{:02d}'.format(num+1) + '</div>',
            html="""<span class="fa-stack " style="font-size: 8pt">
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
data = pd.read_csv(".//loc_name.csv")
print(data)
#'#F0F7EB','#76D8F2','#E2F55B','#F99403','#F0301A','#E50BC4','#B00BF2'
def color(value):
    if (0 < value < 10):
        color = '#76D8F2'
    elif (10 <= value < 25):
        color = '#E2F55B'
    elif (25 <= value < 50):
        color = '#F99403'
    elif (50 <= value < 100):
        color = '#F0301A'
    else:
        color = '#E50BC4'
    return color
for i in range(0,len(data)):
   loc = [data.iloc[i]['lat'], data.iloc[i]['lon']]
   folium.Marker(
        location=loc,
        popup=data.iloc[i]['lat'],
        icon=folium.Icon(icon_color=color(data.iloc[i]['lat']),icon='location-pin'),
        markerColor='black',
    ).add_to(m)
   #folium.Marker(
   #     location=loc,
   #     popup=data.iloc[i]['lat'],
   #     icon= number_DivIcon(color(data.iloc[i]['lon']),0),
   # ).add_to(m)
def add_categorical_legend(folium_map, title, colors, labels):
    if len(colors) != len(labels):
        raise ValueError("colors and labels must have the same length.")

    color_by_label = dict(zip(labels, colors))
    
    legend_categories = ""     
    for label, color in color_by_label.items():
        legend_categories += f"<li><span style='background:{color}'></span>{label}</li>"
        
    legend_html = f"""
    <div id='maplegend' class='maplegend'>
      <div class='legend-title'>{title}</div>
      <div class='legend-scale'>
        <ul class='legend-labels'>
        {legend_categories}
        </ul>
      </div>
    </div>
    """
    script = f"""
        <script type="text/javascript">
        var oneTimeExecution = (function() {{
                    var executed = false;
                    return function() {{
                        if (!executed) {{
                             var checkExist = setInterval(function() {{
                                       if ((document.getElementsByClassName('leaflet-top leaflet-right').length) || (!executed)) {{
                                          document.getElementsByClassName('leaflet-top leaflet-right')[0].style.display = "flex"
                                          document.getElementsByClassName('leaflet-top leaflet-right')[0].style.flexDirection = "column"
                                          document.getElementsByClassName('leaflet-top leaflet-right')[0].innerHTML += `{legend_html}`;
                                          clearInterval(checkExist);
                                          executed = true;
                                       }}
                                    }}, 100);
                        }}
                    }};
                }})();
        oneTimeExecution()
        </script>
      """
   

    css = """

    <style type='text/css'>
      .maplegend {
        z-index:9999;
        float:right;
        background-color: rgba(255, 255, 255, 1);
        border-radius: 5px;
        border: 2px solid #bbb;
        padding: 10px;
        font-size:12px;
        positon: relative;
      }
      .maplegend .legend-title {
        text-align: left;
        margin-bottom: 5px;
        font-weight: bold;
        font-size: 90%;
        }
      .maplegend .legend-scale ul {
        margin: 0;
        margin-bottom: 5px;
        padding: 0;
        float: left;
        list-style: none;
        }
      .maplegend .legend-scale ul li {
        font-size: 80%;
        list-style: none;
        margin-left: 0;
        line-height: 18px;
        margin-bottom: 2px;
        }
      .maplegend ul.legend-labels li span {
        display: block;
        float: left;
        height: 16px;
        width: 30px;
        margin-right: 5px;
        margin-left: 0;
        border: 0px solid #ccc;
        }
      .maplegend .legend-source {
        font-size: 80%;
        color: #777;
        clear: both;
        }
      .maplegend a {
        color: #777;
        }
    </style>
    """

    folium_map.get_root().header.add_child(folium.Element(script + css))

    return folium_map
labels0 = ['Không mưa','0-16 mm','15-25 mm','25-50 mm','50-100 mm','100-200 mm','> 200 mm']
color_bar = ['#F0F7EB','#76D8F2','#E2F55B','#F99403','#F0301A','#E50BC4','#B00BF2']
m = add_categorical_legend(m, 'Chú giải',
                        colors = color_bar,
                        labels = labels0)
   #folium.Marker(
   #   location=[data.iloc[i]['lat'], data.iloc[i]['lon']],
   #   popup=data.iloc[i]['Station'],icon=folium.DivIcon(html=(
   # '<svg height="100" width="100">'
   # '<circle cx="50" cy="50" r="40" stroke="yellow" stroke-width="3" fill="none" />'
   # '<text x="150" y="30" fill="black">{:.1f}</text>'
   # '</svg>'.format(50)
   #).add_to(m)
m.save(html_file)
import webview
# define an instance of tkinter
#  size of the window where we show our website
# Open website
webview.create_window('Bando',url=html_file,width=1024,height=768)
webview.start()