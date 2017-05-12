import pandas
import numpy as np
import matplotlib
from openpyxl import Workbook, worksheet, load_workbook
from os import listdir
from os.path import isfile, join

mypath = 'C:/Python34/Programs/Logs'
onlyfiles = [f for f in listdir(mypath) if isfile(join(mypath, f))]

df = pandas.DataFrame()
for i in onlyfiles:
    excel_file = pandas.ExcelFile('C:/Python34/Programs/logs/'+i)
    df2 = excel_file.parse('Sheet1')
    df2['month']=i[:i.index('.')]
    df = df.append(df2)



df = df.fillna(value = 0)
for i in range(1,38):
    try:
        df.drop('Unnamed: '+str(i), axis = 1, inplace = True)
    except ValueError:
        pass

def totalByMonth(value, month):
    """ Takes in two strings and returns sum of column
        matching the value string where month column
        matches given month string.
    """
    while value not in df.columns:
        print('invalid value')
        value = input('enter a new value: ')
    return(np.sum(df[value][df['month']==month]))

def allTotalsByMonth(month):
    """Takes input as a string and returns the sum of all columns
       except for 'name' and 'month' columns. Returns a list of summed
       values.
       Returns a dict.
    """
    totals_dict = {}
    for i in df.columns[1:-1]:
        amount = df.loc[df['month']==month, i].sum()
        totals_dict[i]=amount
    final = sortdictionary(totals_dict)
    return final

def sortdictionary(dictionary):
    """Receives a dictionary as input. Sorts the dictionary based on key values
       in ascending order into a list. Then reverses that list so that values are
       in descending order (highest values first).
       Returns a list.
    """
    dictionary = sorted(dictionary, key=dictionary.__getitem__)
    results = []
    for i in dictionary:
        value = i, dictionary[i]
        results.append(value)
    results.reverse()
    return results
print(allTotalsByMonth('november'))

def totalForAllMonths(df):
    """Inputs a dataframe and returns sum of all rows for each
        column of problems.
        Returns a dict.
    """
    totals = {}
    for i in df.columns[1:-1]:
        totals[i] = np.sum(df[i])
    return totals

def mostFrequentProblemAll():
    """ Takes in a month as a string. Creates a list of column names from
        dataframe. Sets a variable to hold highest value to Zero. Iterates
        through index with only problems and sums the values in each column.
        Returns tuple with of (problem, amount).
    """
    skipcolumns = ['"Male = M/Female = F"', 'Number of Contacts',
                 'Open Frnotier Health Case', 'Collaborative Contacts',
                 'month', 'New Student', 'Group Therapy']
    column_name = list(df.columns)
    most = 0
    for c in column_name[7:-1]:
        if c in skipcolumns:
            pass
        else:
            amount = (np.sum(df[c]))
            if amount > most:
                most = amount
                name = c
    return name, most

def mostFrequentProblembyMonth(month):
    """Inputs a month. Calls function to return dict of summed values
        for that month. Iterates over the keys in that dict and checks they are
        problem keys. Then compares the value of that key with a variable
        holding the highest found value. That variable (most) is initially
        set to zero. Returns key and value of highest value in dictionary.
    """
    totals = allTotalsByMonth(month)
    skipcolumns = ['"Male = M/Female = F"', 'Number of Contacts',
                 'Open Frnotier Health Case', 'Collaborative Contacts',
                 'month', 'New Student', 'Group Therapy']
    most = 0
    for k in totals:
        if k in skipcolumns:
            pass
        else:
            value = totals[k]
            if value > most:
                most = value
                key = k
    return key, totals[key]

def searchByCritirea(crit1, crit2, crit2value):
    """Recieves problem to search for and returns that problem
       when value of crit2 matches value given to function as crit2value.
       Returns int.
    """
    result = df.loc[df[crit2] == crit2value, crit1].sum()
    return result

def averageContactByVariable(variable):
    """Recieve a problem to search in. Calculates sum of column and divides by
       number of non-zero values in that column.
    """
    amount = np.sum(df[variable])
    count = np.sum(df[variable] != 0)
    avg = amount/count
    return avg

def percentByCritirea(crit1, crit2, crit2value):
    """Receives problem to search for and criteria name to search by and critirea name
        value to match. Calculates percent
        and returns as an int.
        ex. percentByCritirea('Depression', '"Male = M/Female = F"', 'm') returns percent
        of male contacts seen for depression.
    """
    contacts = searchByCritirea('Number of Contacts', crit2, crit2value)
    problemcontacts = searchByCritirea(crit1, crit2, crit2value)
    percent = (problemcontacts*100)/contacts
    return percent

