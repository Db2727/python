from random import randint, seed
seed()


iterations = 1000
num = 0


def roll():
    w1 = randint(1, 6)
    w2 = randint(1, 6)
    z = 1
    while w1 != w2:
        w1 = randint(1, 6)
        w2 = randint(1, 6)
        z = z + 1
    return z

for i in range(iterations):
    z = roll()
    num = num + z
print(f"Durchschnitt: {num / iterations}")