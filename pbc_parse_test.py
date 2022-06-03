org_f = open("pbc_parse_test.txt", "r")

data = org_f.readlines()

for item in data:
    targ = item.split(':')[2].split(' ')[0]
    print(targ)
