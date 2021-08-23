import cmath
import math
import numpy as np
import os
import openpyxl
import xlrd
import pyexcel as p
infile = open(r'\\gcsadfs/share/public/DMohata/GaN HEMT/Data/Rockling/sparameters/948_dev5_28V_vgssweep.dat', 'r')
for line in infile:
    # Typical line: variable = value
    variable, value = line.split('=')
    variable = variable.strip()  # remove leading/traling blanks
    if variable == 'v0':
        v0 = float(value)
    elif variable == 'a':
        a = float(value)
    elif variable == 'dt':
        dt = float(value)
    elif variable == 'interval':
        interval = eval(value)
infile.close()


#try:
 #   data=np.loadtxt(r'\\gcsadfs/share/public/DMohata/GaN HEMT/Data/Rockling/sparameters/948_dev5_28V_vgssweep.dat', dtype='float')
#except ValueError:
 #   continue


x=1
y=1

z = complex(x, y);

# converting complex number into polar using polar()
w = cmath.polar(z)

# printing modulus and argument of polar complex number
print("The modulus and argument of polar complex number is : ", end="")
print(w)

# converting complex number into rectangular using rect()
w = cmath.rect(1.4142135623730951, 0.7853981633974483)

# printing rectangular form of complex number
print("The rectangular form of complex number is : ", end="")
print(w)