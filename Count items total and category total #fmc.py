# Nested Dictionaries and Lists
# useful to count items in a total menu and the total
"""
This may seem like such a simple thing to model that you wouldn’t
need to bother with writing a program to do it. But realize that this same
totalBrought() function could easily handle a dictionary that contains
thousands of guests, each bringing thousands of different picnic items. Then
having this information in a data structure along with the totalBrought()
function would save you a lot of time!
You can model things with data structures in whatever way you like, as
long as the rest of the code in your program can work with the data model
correctly. When you first begin programming, don’t worry so much about
the “right” way to model data. As you gain more experience, you may come
up with more efficient models, but the important thing is that the data
model works for your program’s needs.

#fmc
"""
allGuests = {'alice': {'apple':  1, 'orange': 2},
            'bob':   {'orange': 3, 'carrot': 4},
            'mike':  {'carrot': 1, 'tomato': 5}}


def totalBrought(guests, item):
    numBrought = 0
    for k, v in guests.items():
        numBrought += v.get(item, 0)
    return numBrought


print('Number of things being brought:')
print(' - apple ' + str(totalBrought(allGuests, 'apple')))
print(' - orange ' + str(totalBrought(allGuests, 'orange')))
print(' - carrot ' + str(totalBrought(allGuests, 'carrot')))
print(' - tomato ' + str(totalBrought(allGuests, 'tomato')))
print('TOTAL: ' + str(totalBrought(allGuests, 'apple') + totalBrought(allGuests, 'orange') + \
totalBrought(allGuests, 'carrot') + totalBrought(allGuests, 'tomato')))












