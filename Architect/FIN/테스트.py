a=1
b=2
c=3


print(f"{a}{b}{c}")

print(f'{float(a):0.3f}')


print(f"{float(a):0.3f}")


width = 200
height = 5


standard = f'{format(float(width), ".3f")}*{format(float(height), ".3f")}={format(float(width)*float(height), ".3f")}'

standard2 = f'{float(width):0.3f}*{float(height):0.3f}={float(width)*float(height):0.3f}'

print(standard, '/', standard2)



