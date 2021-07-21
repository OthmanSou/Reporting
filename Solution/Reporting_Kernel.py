# -*- coding: utf-8 -*-
"""
Created on Wed Jul 21 15:40:15 2021

@author: osoua
"""

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


param_data_frame = pd.read_excel(File_name,sheet_name = 'Main', skiprows=5, usecols='G:I')
Report_number = len(param_data_frame)

time_series = {}
reportingDF = {}
n = 20

for Pos in range(Report_number):
	Start_date = "2019-01-01T09:00:00"
	End_date = param_data_frame.iloc[Pos]['Analysis Date'].strftime('%Y-%m-%d')
	Underlying = [param_data_frame.iloc[Pos]['Underlyings']]
	Interval = param_data_frame.iloc[Pos]['Time Frame']
	
	#time_series[Underlying[0]] = pd.read_excel(r'C:\Users\osoua\Desktop\Work\Reporting Financier\Solution\RFR Lib6M.xlsx')
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
	time_series[Underlying[0]]['H-C'] = time_series[Underlying[0]]['HIGH']-time_series[Underlying[0]]['CLOSE']
	time_series[Underlying[0]]['H-PC'] = time_series[Underlying[0]]['HIGH']-time_series[Underlying[0]]['P_Close']
	time_series[Underlying[0]]['PC-L'] = time_series[Underlying[0]]['P_Close']-time_series[Underlying[0]]['LOW']
	time_series[Underlying[0]]['TR'] = time_series[Underlying[0]][['H-C','H-PC','PC-L']].max(axis = 1)
	time_series[Underlying[0]].drop(['P_Close','H-C','H-PC','PC-L'], axis =1, inplace=True)
	#Average True Range, or VOLK ATR = TR's exponential moving average
	time_series[Underlying[0]]['VOLK'] = time_series[Underlying[0]]['TR'].ewm(span=n, adjust=False).mean()
	time_series[Underlying[0]].drop(['TR'], axis =1, inplace=True)
	
	#Import the reporting DataFrame format and fill It
	reportingDF[Underlying[0]] = pd.read_excel(File_name, sheet_name = 'Reporting', skiprows=5, usecols='F:K')
	#Semaine 
	
	#Diff Hebdo
	#VAR 01/01/2021
	
	#VAR MTD
	
	#VAR YTD
	
	
#Max Min YTD MTD YTY WTD VAR

#Display


	
	
	
 
	
