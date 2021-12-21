import re

def find_nearest(items, pivot):
    return min(items, key = lambda x:abs(x - pivot ))

def generate_parameters():
    with open('parameters.txt') as file:
        lines = file.readlines()
        lines = [line.rstrip() for line in lines]

    return(lines)

# key for natural sorting
# https://stackoverflow.com/questions/4836710/is-there-a-built-in-function-for-string-natural-sort
natsort = lambda s: [int(t) if t.isdigit() else t.lower() for t in re.split('(\d+)', s)]