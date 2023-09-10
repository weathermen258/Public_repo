import fileinput
from tkinter import Tk
from tkinter.filedialog import askopenfilename
# Does a list of files, and
# redirects STDOUT to the file in question
Tk().withdraw() # disable the root windows
input_file = askopenfilename() # open file browsing dialog box
rt = open(input_file,'r+')
next(rt)
print(rt)
for i in range(len(rt)):
   check_line = f'{stations[i]},{obs_year},{obs_month},{obs_day},{obs_hour},{obs_min}'
   print(check_line)
   for line1 in rt:
      if check_line in line1:
         replace_line="{},{},{},{},{},{},{},{},{},{},{},{}\n"\
                    .format("stations","obs_year","obs_month","obs_day",\
                        "obs_hour","obs_min","wind_dir","wind_speed",\
                        "slp","delta_p","gust_dir","gust_speed")
         print(replace_line)
         rt.write(line1.replace(line1,replace_line))
      else:
         rt.write(write_line)

