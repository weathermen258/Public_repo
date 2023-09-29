def test1():
    global a,b,c
    a = []
    b = []
    c = []
    def make_data():
        for i in range(10):
            a.append(i)
            b.append(i+1)
            c.append(i*2)
    make_data()
    return a,b,c
test1()
print (a)
print (b)
print (c)

