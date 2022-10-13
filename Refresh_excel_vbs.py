#!/usr/bin/env python
# coding: utf-8

# In[26]:


import os
import sys
import subprocess
import shutil
import win32com.client

File = win32com.client.DispatchEx("Excel.Application")
File.Visible = True
src = r'C:\Users\wendysu\OneDrive - Micron Technology, Inc\GDM\GDM_dashboard\Run_Task'

for dirpath,dirnames, filenames in os.walk(src):
    for filename in filenames:
        file = os.path.join(dirpath,filename)
        subprocess.call(['cscript.exe', file])

File.Quit()


# In[ ]:




