#!/usr/bin/env python
# coding: utf-8

# In[1]:


import numpy as np
import pandas as pd
from docx import Document


# # Functions

# In[2]:


def reformat_date(date_str):
    """reformat python datetime to required string format"""
    month, day, time = date_str.split('-')
    return f"{int(month)}月{int(day)}日/ {int(time)}点"


# # Read DataFrame

# In[3]:


"""read table"""
df = pd.read_excel('计划汇集报送.xlsx', header=1)


# # Feature engineering

# In[4]:


# Reformat date
df['计划日期'] = df['计划日期'].dt.strftime('%m-%d-%H')
df['计划日期'] = df['计划日期'].apply(reformat_date)


# In[5]:


doc = Document()

for index, row in df.iterrows():
    platform_name = row['到站名称']
    station_name = row['液源']
    planning_datetime = row['计划日期']
    customer_name = row['计划所属客户名称']
    car_number = row['车牌号/挂车号']
    driver_info = row['司机姓名/电话']
    supercargo_info = row['押运员姓名/电话']
    
    formatted_paragraph = (f"{platform_name}（{station_name}） {planning_datetime}-{customer_name}\n"
    f"车牌号：{car_number}\n"
    f"驾驶员：{driver_info}\n"
    f"押运员：{supercargo_info}")
    
    doc.add_paragraph(formatted_paragraph)


# In[6]:


doc.save('test output.docx')


# In[ ]:




