import random
import base64

from getmac import get_mac_address

print("M---", get_mac_address())
print("E---", base64.b64encode(get_mac_address().encode('ascii')))
f = open("moonlight.dat", "wb")
f.write(base64.b64encode(get_mac_address().encode('ascii')))
for x in range(2**10) :
    print(random.random(), random.gauss(10, 3))
f.close()

#--onefile