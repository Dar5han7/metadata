##logic is not clear###
n=int(input("no of test"))
l = {}
for i in range(0,n):
    j = int(input("enter the patients"))
    k = int(input("enter the interval"))
    l[i]= [j,k]
w =[]
for id, test in l.items():

    for c,p in enumerate(test):
        if c ==0:
            j=p
        else:
            k=p

    w.append((j-1)*k)
for x in w:
    print(x)



