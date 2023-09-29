def pos(str1,str2):
   pos = -1
   if str1 in str2:
      for i in range(len(str2.split())):
         if str1 == str2.split()[i]:
            pos = i
   return pos
line1='48/66 32498 73602 10284 20261 30020 40025 53018 \
879// 333 59038 83894 82995 86697='
print(line1.split())
#print(pos('66',line1))
if pos('59038',line1) == pos('333',line1) + 1:
   print('good its true')
else:
   print('its bad')
