#sheet
from pyparsing import col


def column(i):
    '''0=A
    25=z
    26=az'''

    a = ord('A')

    i-=1

    if i + a > ord("Z"):
        return "" + chr((i//26) + ord("A") - 1) + chr((i%26) + ord("A"))

    return chr(i + a)

print(column(26*3))
