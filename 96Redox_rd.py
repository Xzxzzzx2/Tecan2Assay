from csv import reader
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import openpyxl

file_path_ox = 'TFP p2 ox.xlsx' #import an excel file with 2 sheets
file_path_rd = 'TFP p2 rd.xlsx' #import an excel file with 2 sheets
title_index = 46 #change to row number of the title
last_index = 68 #change to row number of the last line of data
sheet1 = 'ox' #change to oxidation sheet name
sheet2 = 'rd' #change to reduction sheet name

#raw_df = pd.ExcelFile(file_path)
#raw_df1 = pd.read_excel(raw_df, sheet1)
#raw_df2 = pd.read_excel(raw_df, sheet2)
raw_df1 = pd.read_excel(file_path_ox)
raw_df2 = pd.read_excel(file_path_rd)
title_index = title_index - 2
last_index = last_index -1

raw_list1 = raw_df1.values.tolist()
raw_list2 = raw_df2.values.tolist()

clean_list1 = []
clean_list2 = []
for i in list(range(title_index, last_index)):
    clean_list1.append(raw_list1[i])
    clean_list2.append(raw_list2[i])
    
ox = pd.DataFrame(clean_list1[1:], columns = clean_list1[0])
rd = pd.DataFrame(clean_list2[1:], columns = clean_list2[0])
print(ox)
print(rd)

fig, axs = plt.subplots(8, 6, figsize=(36, 36))

indices = [(0,0), (1,0), (2,0), (3,0), (4,0), (5,0), (6,0), (7,0),
           (0,1), (1,1), (2,1), (3,1), (4,1), (5,1), (6,1), (7,1),
           (0,2), (1,2), (2,2), (3,2), (4,2), (5,2), (6,2), (7,2),
           (0,3), (1,3), (2,3), (3,3), (4,3), (5,3), (6,3), (7,3),
           (0,4), (1,4), (2,4), (3,4), (4,4), (5,4), (6,4), (7,4),
           (0,5), (1,5), (2,5), (3,5), (4,5), (5,5), (6,5), (7,5)]

t1 = ['A7', 'B7', 'C7', 'D7', 'E7', 'F7', 'G7', 'H7',
      'A8', 'B8', 'C8', 'D8', 'E8', 'F8', 'G8', 'H8',
      'A9', 'B9', 'C9', 'D9', 'E9', 'F9', 'G9', 'H9',
      'A10', 'B10', 'C10', 'D10', 'E10', 'F10', 'G10', 'H10',
      'A11', 'B11', 'C11', 'D11', 'E11', 'F11', 'G11', 'H11',
      'A12', 'B12', 'C12', 'D12', 'E12', 'F12', 'G12', 'H12']

for i in range(48):
    if t1[i] in list(ox): 
        axs[indices[i]].plot(wavelen, ox[t1[i]], label = "Oxidation", color= "orange")
        axs[indices[i]].plot(wavelen, rd[t1[i]], label = "Reduction", color= "blue")
        axs[indices[i]].set_title(t1[i])
        
#fig.patch.set_facecolor('xkcd:beige')
fig.patch.set_facecolor('xkcd:white')

plt.show()
