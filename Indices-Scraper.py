#!/usr/bin/env python
# coding: utf-8

# In[1]:


# import libraries
import requests
from bs4 import BeautifulSoup
import pandas as pd
from textwrap import wrap
import re
import jdatetime


# In[2]:


# URLs
file_path = './data/Industry-Indices.xlsx'
tse_main_url = 'http://www.tsetmc.com/Loader.aspx?ParTree=15'
industries_operation_url = 'http://www.tsetmc.com/Loader.aspx?Partree=15131O'


# ## Fetching data

# In[4]:


def get_and_parse_url(url, params=None):
    """
        get a URL, grab the data, and return a BeautifulSoup object with parsed data
    """
    res = requests.get(url=url, params=params)
    soup = BeautifulSoup(markup=res.text, features='html.parser')
    return soup


# In[5]:

print('Fetching requested data from website ...')
soup = get_and_parse_url(url=industries_operation_url)
if soup:
    print('Fetching done.')


# In[6]:

print('Finding out the date')
msoup = get_and_parse_url(url=tse_main_url)

# ## Find the date (Jalali) and convert it to Gregorian

# In[7]:


# Market info located in a blue div
blue_div = msoup.find_all(name='div', class_='box1 blue tbl z1_4 h210')


# In[9]:

# Grab the datetime out of the blue div
tds = blue_div[0].find_all('td')
info_datetime = tds[9].string
info_datetime = '13'+ info_datetime


# In[10]:

# throw out the time part and clean the date part
info_j_date = info_datetime[:10]
info_j_date = info_j_date.strip().split('/')
info_j_date[1] = ('0' + info_j_date[1]) if len(info_j_date[1]) < 2 else info_j_date[1]
info_j_date[2] = ('0' + info_j_date[2]) if len(info_j_date[2]) < 2 else info_j_date[2]


# In[11]:


# convert Jalali date to Gregorian date
info_c_date = jdatetime.date(year=int(info_j_date[0]), 
                             month=int(info_j_date[1]), 
                             day=int(info_j_date[2])).togregorian()


# In[12]:


# convert both Jalali and Gregorian date to string
info_c_date = info_c_date.strftime('%Y-%m-%d')
info_j_date = '{}-{}-{}'.format(info_j_date[0], info_j_date[1], info_j_date[2])
print('Date is:')
print(info_c_date)
print(info_j_date)


# ## Searching for required data and preparing them

# In[13]:
print('Searching for required data and preparing them...')
rows = soup.tbody.find_all(name='tr')




# In[15]:


rows = [r.find_all('td') for r in rows]


# In[16]:


# col1: Group
# col2: Market-Value
# col3: Transactions-Number
# col4: Transactions-Volume
# col5: Transactions-Value


# In[17]:


# get cells string and save them in a list
values = []
for row in rows:
    for col in row:
        values.append(col.string)
        


# In[18]:


# split the values into sized 5 chunks to represent each row in a list item
values = [values[i:i+5] for i in range(0, len(values), 5)]


# In[19]:


# remove the ',', 'B' and 'M' from the recieved string and convert it to a float number
def purify_number(number):
    number = str(number)
    number = number.split(',')
    number = ''.join(number)
    number = number.strip()

    if 'B' in number:
        number = number.strip('B')
        number = float(number) * 1000
    elif 'M' in number:
        number = number.strip('M')
        #number = float(number) * 1000000
    else:
        number = float(number)

    return number



# In[21]:


# cleaning all the numbers
for i in range(len(values)):
    for j in range(1, len(values[0])):
        values[i][j] = purify_number(values[i][j])


# In[22]:


# now data is clean and ready to save
print('Data have been cleaned and ready to save.')

# ## Saving the clean data into excel file

# In[23]:


# read the original excel file
df_main = pd.read_excel(file_path)
df_main.head()


# In[24]:


# find the group number using regex
def parse_group_no(text):
    if re.search('\d+', text):
        group_no = re.findall('\d+', text)[0]
    else:
        group_no = '0'
        
    return group_no


# In[25]:


# prepare a dict of values for creating a DataFrame
CDate = info_c_date * len(values)
CDate = wrap(text=CDate, width=10)
JDate = info_j_date * len(values)
JDate = wrap(text=JDate, width=10)
data = {
    'CDate': CDate,
    'JDate': JDate,
    'GroupNo': [parse_group_no(values[i][0]) for i in range(len(values))],
    'GroupName': [(values[i][0]).encode('utf-8') for i in range(len(values))],
    'MarketValue': [values[i][1] for i in range(len(values))],
    'TransactionsCount': [values[i][2] for i in range(len(values))],
    'TransactionsVol': [values[i][3] for i in range(len(values))],
    'TransactionsValue': [values[i][4] for i in range(len(values))]
}


# In[27]:


# create new DataFrame with recently fetched data
df = pd.DataFrame(data=data)


# In[28]:


# read the original excel file
print('Opening the Excel file and saving data...')
df_main = pd.read_excel(file_path)


# In[29]:


# combine recently created DataFrame and the original one
df1 = pd.concat([df_main, df], ignore_index=True, sort=False)


# In[30]:


# Save the output to excel file
df1.to_excel(excel_writer=file_path, index=False)
print('Data stored in the Excel file successfully.')