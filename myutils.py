import re

def find_nearest(items, pivot):
    return min(items, key = lambda x:abs(x - pivot ))

def generate_parameters(filename):
    with open(filename) as file:
        lines = file.readlines()
        lines = [line.rstrip() for line in lines]

    return(lines)

# key for natural sorting
# https://stackoverflow.com/questions/4836710/is-there-a-built-in-function-for-string-natural-sort
natsort = lambda s: [int(t) if t.isdigit() else t.lower() for t in re.split('(\d+)', s)]

def replace_list_elements_by_dict(mylist, mydict):
    # Acts in place
    for i in range(len(mylist)):
        if mylist[i] in mydict.keys():
            mylist[i] = mydict[mylist[i]]

def dict_from_two_lists(list_keys, list_values):
    return dict(zip(list_keys, list_values))

def orange(text):
    COLOR = '\033[93m'
    END = '\033[0m'
    return f'{COLOR}{text}{END}'

def red(text):
    COLOR = '\033[91m'
    END = '\033[0m'
    return f'{COLOR}{text}{END}'