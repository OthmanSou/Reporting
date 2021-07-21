# -*- coding: utf-8 -*-
"""
Reporting v 0.0
"""

import eikon as ek
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

#variables and arguments
App_key = r'310a8e0605ee4dbe9befd54b7911c060428878ae'
File_name = r'C:\Users\osoua\Desktop\Work\Reporting Financier\Solution\Reporting.xlsm'


ek.set_app_key(App_key)


param_data_frame = pd.read_excel(File_name, skiprows=5, usecols='G:I')
Report_number = len(param_data_frame)

time_series = {}
n = 20

for Pos in range(Report_number):
	Start_date = "2019-01-01T09:00:00"
	End_date = param_data_frame.iloc[Pos]['Analysis Date'].strftime('%Y-%m-%d')
	Underlying = [param_data_frame.iloc[Pos]['Underlyings']]
	Interval = param_data_frame.iloc[Pos]['Time Frame']
	time_series[Underlying[0]] = ek.get_timeseries(Underlying, 
									start_date = Start_date,
								    end_date = End_date, 
									interval = Interval)
	#Typical Price, TP = (High + Low + Close) /3
	time_series[Underlying[0]]['TP'] = (time_series[Underlying[0]]['HIGH'] +
										time_series[Underlying[0]]['LOW']
										+ time_series[Underlying[0]]['CLOSE'])/3
	#VOL Bollinger
	time_series[Underlying[0]]['VOLB'] = time_series[Underlying[0]]['TP'].rolling(n).std()
	#Previous Close
	time_series[Underlying[0]]['P_Close'] = time_series[Underlying[0]]['CLOSE'].shift(periods=1)
	#True Range, TR = max(High-Low, High-PreviousClose, PreviousClose-Low)
	#time_series[Underlying[0]]['TR'] = pd.DataFrame(time_series[Underlying[0]]['HIGH']-time_series[Underlying[0]]['CLOSE'],
										#time_series[Underlying[0]]['HIGH']-time_series[Underlying[0]]['LOW']).max(axis = 1)
	#exponential moving average
	
	#VOL K
	
	#Vol B
	
	#Diff Hebdo
	
	#VAR 01/01/2021
	
	#VAR MTD
	
	#VAR YTD
	
	
#Max Min YTD MTD YTY WTD VAR

#Display


	
	
	
 
	
