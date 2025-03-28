string = "1. Firmenname"

def split_name(name):
    x = name.split(".", 1)
    return int(x[0])

print(split_name(string))