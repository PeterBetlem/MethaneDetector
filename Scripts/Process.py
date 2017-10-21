# -*- coding: utf-8 -*-
"""
Created on Sat Oct 21 20:35:07 2017

@author: Peter
"""
#==============================================================================
# Import required libraries
#==============================================================================
import numpy as np
import csv
import pandas as pd
import matplotlib.pyplot as plt
from Tkinter import Tk
from tkFileDialog import askopenfilename
#import xlsxwriter

#==============================================================================
# Setting variables
#==============================================================================
TimeBetweenMeasurements = 5 #Minutes between datasets...
MinimalValue = 1

#==============================================================================
# Selecting File using GUI
#==============================================================================
#Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
#filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file
filename = '../Data/20172021_lagoon.log'
print(filename)

#==============================================================================
# Scanning file for data, and loading data
#==============================================================================

Phrase1 = 'File Contents:'
Phrase2 = 'Finished.'
with open(filename) as input_data:
    for num, line in enumerate(input_data, 1):
        if Phrase1 in line:
#            print 'Data starts at line:', num+2
            start = num
            break
    for num, line in enumerate(input_data, 1):        
        if Phrase2 in line:
#            print 'Data ends at line:', num-2
            end = num-2
            break

data = pd.read_csv(filename,header=0,skiprows=start,nrows=end-start,delimiter=',',parse_dates=True)
#==============================================================================
# Removing bogus data
#==============================================================================
data = data.drop(data.columns[[0]], axis=1)  
data = data.drop(data.index[0]) 

#==============================================================================
# Reorganising based on timestamp
#==============================================================================
data['datetime'] = pd.to_datetime(data['Year'] + data['Month'] + data['Day'] + data['Hour'] + data['Minute'] + data['Second'])
data = data.set_index(pd.DatetimeIndex(data['datetime']))
cols = data.columns.tolist()
cols = cols[6:-1]
data = data[cols]

#==============================================================================
# Converting all columns to numeric data
#==============================================================================
data = data.apply(pd.to_numeric)

#==============================================================================
# Calculating time difference between each datapoint in minutes
#==============================================================================
data['deltaT'] = data.index.to_series().diff().dt.seconds.div(60, fill_value=0)

#==============================================================================
# Sorting measurements based on acceptable timedifference
#==============================================================================
Measurements = data.index[data['deltaT'] > TimeBetweenMeasurements]

#==============================================================================
# Outputting each dataset
#==============================================================================
for x in xrange(len(Measurements)):
    if x == 0:
        dataset = data.loc[data.index[0]:Measurements[x]]
    elif x == len(Measurements):
        dataset = data.loc[data.index[x]:data.index[len(data)-1]]
    else:
        dataset = data.loc[Measurements[x-1]:Measurements[x]]

    processed = dataset[dataset['CH4 (PPM)'] >= MinimalValue]
    
    cols = processed.columns.tolist()
    cols = cols[2:6]
    processed = processed[cols]  

    if len(processed) >= 10:
        plt.figure(); processed['CH4 (PPM)'].plot()
#==============================================================================
#     Outputting to excel
#==============================================================================
        filename_folder = os.path.dirname(os.path.abspath(filename))
        name = str(processed.index[0].strftime('%Y-%m-%d %H%M%S'))+'.xlsx'
        name = os.path.join(filename_folder, name)
        writer = pd.ExcelWriter(name, engine='xlsxwriter')
        processed.to_excel(writer,sheet_name='Processed data')
        dataset.to_excel(writer,sheet_name='Full dataset')
#        writer.save()
#==============================================================================
#       Adding a chart to excel to display the data
#==============================================================================
        workbook  = writer.book
        worksheet = writer.sheets['Processed data']

#==============================================================================
#       Setting Chart type
#==============================================================================
        chart = workbook.add_chart({'type': 'scatter',
                             'subtype': 'smooth_with_markers'}) #Also try: smooth
#==============================================================================
#       Setting Chart Title, axes, and ticks
#==============================================================================
        chart.set_title ({'name': 'CH4 (ppm) Raw Values'})
        chart.set_x_axis({'name': 'Measurement'})
        chart.set_y_axis({'name': 'CH4 (ppm)'})
        chart.set_legend({'none': True})
        chart.set_x_axis({'date_axis': True,'num_font':  {'rotation': 0},'num_format': 'h:mm'})
#==============================================================================
#       Add data to be plotted
#==============================================================================
        chart.add_series({
                'categories': ['Processed data', 1,0,len(dataset)+1,0],
                'values': ['Processed data', 1,1,len(dataset)+1,1],
                })
#==============================================================================
#       Insert the chart onto the sheet
#==============================================================================
        worksheet.insert_chart('G1',chart)
#==============================================================================
#       Save the excel document in the output folder
#==============================================================================
        writer.save()

