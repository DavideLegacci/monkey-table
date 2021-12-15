with open('parameters.txt') as file:
    lines = file.readlines()
    lines = [line.rstrip() for line in lines]

print(lines)