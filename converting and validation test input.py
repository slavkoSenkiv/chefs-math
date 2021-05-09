# Will be useful if user inputs different variations of the dish's name
"""
for exsample the recepie is 'гороховий суп' and the user will input 'Гороховий суп'

The upper() and lower() methods are helpful if you need to make a case insensitive comparison. For example, the strings 'great' and 'GREat' are not
equal to each other. But in the following small program, it does not matter
whether the user types Great, GREAT, or grEAT, because the string is first converted to lowercase.
"""

print('how are you')
feeling = input()
if feeling.lower() == 'great':
    print('+')
else:
    print('-')
    