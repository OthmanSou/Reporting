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
import refinitiv.dataplatform as rdp 
import numpy as np
import matplotlib.pyplot as plt
import QuantLib as ql
from datetime import datetime as dt

def ql_to_string(d):
	month = str(d.month())
	if d.month()<10: 
		month ='0' + month
	day = str(d.dayOfMonth())
	if d.dayOfMonth()<10:
		day ='0'+ day
	return str(d.year()) + '-' + month + '-' + day
    

#variables and arguments
App_key = r'310a8e0605ee4dbe9befd54b7911c060428878ae'
File_name = r'C:\Users\osoua\Desktop\Work\Reporting Financier\Solution\Reporting.xlsm'

ek.set_app_key(App_key)
rdp.open_desktop_session(App_key)


param_data_frame = pd.read_excel(File_name,sheet_name = 'Main', skiprows=5, usecols='G:I')
Report_number = len(param_data_frame)

time_series = {}
reportingDF = {}
n = 20

for Pos in range(Report_number):
	Start_date = "2019-01-01"
	End_date = param_data_frame.iloc[Pos]['Analysis Date'].strftime('%Y-%m-%d')
	Underlying = [param_data_frame.iloc[Pos]['Underlyings']]
	Interval = param_data_frame.iloc[Pos]['Time Frame']
	
	time_series[Underlying[0]] = pd.read_excel(r'C:\Users\osoua\Desktop\Work\Reporting Financier\Solution\TestCAC.xlsx', index_col =0)
	#time_series[Underlying[0]] = ek.get_timeseries(Underlying, 
	#								start_date = Start_date,
	#							    end_date = End_date, 
	#								interval = Interval)
	#Typical Price, TP = (High + Low + Close) /3
	time_series[Underlying[0]]['TP'] = (time_series[Underlying[0]]['HIGH'] +
										time_series[Underlying[0]]['LOW']
										+ time_series[Underlying[0]]['CLOSE'])/3
	#VOL Bollinger
	time_series[Underlying[0]]['VOLB'] = time_series[Underlying[0]]['TP'].rolling(n).std()
	#Previous Close
	time_series[Underlying[0]]['P_Close'] = time_series[Underlying[0]]['CLOSE'].shift(periods=1)
	#True Range, TR = max(High-Low, High-PreviousClose, PreviousClose-Low)
	time_series[Underlying[0]]['H-L'] = time_series[Underlying[0]]['HIGH']-time_series[Underlying[0]]['LOW']
	time_series[Underlying[0]]['H-PC'] = time_series[Underlying[0]]['HIGH']-time_series[Underlying[0]]['P_Close']
	time_series[Underlying[0]]['PC-L'] = time_series[Underlying[0]]['P_Close']-time_series[Underlying[0]]['LOW']
	time_series[Underlying[0]]['TR'] = time_series[Underlying[0]][['H-L','H-PC','PC-L']].max(axis = 1)
	time_series[Underlying[0]].drop(['P_Close','H-L','H-PC','PC-L'], axis =1, inplace=True)
	#Average True Range, or VOLK ATR = TR's exponential moving average
	time_series[Underlying[0]]['VOLK'] = time_series[Underlying[0]]['TR'].ewm(span=n, adjust=False).mean()
	time_series[Underlying[0]].drop(['TR'], axis =1, inplace=True)
	
	#Import the reporting DataFrame format and fill It
	reportingDF[Underlying[0]] = pd.read_excel(File_name, sheet_name = 'Reporting', skiprows=5, usecols='F:J', index_col =0)
	#Semaine 
	reportingDF[Underlying[0]].at['Open','Semaine'] = time_series[Underlying[0]].at[End_date,'OPEN']
	reportingDF[Underlying[0]].at['Close','Semaine'] = time_series[Underlying[0]].at[End_date,'CLOSE']
	reportingDF[Underlying[0]].at['High','Semaine'] = time_series[Underlying[0]].at[End_date,'HIGH']
	reportingDF[Underlying[0]].at['Low','Semaine'] = time_series[Underlying[0]].at[End_date,'LOW']
	reportingDF[Underlying[0]].at['VolK','Semaine'] = time_series[Underlying[0]].at[End_date,'VOLK']
	reportingDF[Underlying[0]].at['VolB','Semaine'] = time_series[Underlying[0]].at[End_date,'VOLB']
	#Diff Hebdo
	Calendar = ql.TARGET()
	Maturity = ql.Period(-1, ql.Weeks)
	Convention = ql.ModifiedFollowing
	QlEnd_date = ql.Date(End_date, '%Y-%m-%d')
	QlPreious_week_date = Calendar.advance(QlEnd_date, Maturity, Convention, False)
	Preious_week_date = ql_to_string(QlPreious_week_date)
	
	reportingDF[Underlying[0]].at['Open','Diff hebdo']  = 100*(time_series[Underlying[0]].at[End_date,'OPEN'] - time_series[Underlying[0]].at[Preious_week_date,'OPEN'])/time_series[Underlying[0]].at[Preious_week_date,'OPEN']
	reportingDF[Underlying[0]].at['Close','Diff hebdo'] = 100*(time_series[Underlying[0]].at[End_date,'CLOSE'] - time_series[Underlying[0]].at[Preious_week_date,'CLOSE'])/time_series[Underlying[0]].at[Preious_week_date,'CLOSE']
	reportingDF[Underlying[0]].at['High','Diff hebdo']  = 100*(time_series[Underlying[0]].at[End_date,'HIGH'] - time_series[Underlying[0]].at[Preious_week_date,'HIGH'])/time_series[Underlying[0]].at[Preious_week_date,'HIGH']
	reportingDF[Underlying[0]].at['Low','Diff hebdo']   = 100*(time_series[Underlying[0]].at[End_date,'LOW'] - time_series[Underlying[0]].at[Preious_week_date,'LOW'])/time_series[Underlying[0]].at[Preious_week_date,'LOW']
	reportingDF[Underlying[0]].at['VolK','Diff hebdo']  = 100*(time_series[Underlying[0]].at[End_date,'VOLK'] - time_series[Underlying[0]].at[Preious_week_date,'VOLK'])/time_series[Underlying[0]].at[Preious_week_date,'VOLK']
	reportingDF[Underlying[0]].at['VolB','Diff hebdo']  = 100*(time_series[Underlying[0]].at[End_date,'VOLB'] - time_series[Underlying[0]].at[Preious_week_date,'VOLB'])/time_series[Underlying[0]].at[Preious_week_date,'VOLB']
	#VAR MTD
	Maturity = ql.Period(0, ql.Days)
	QlFirstOfMonth_date_adjusted = Calendar.advance(ql.Date(1,QlEnd_date.month(),QlEnd_date.year()), Maturity, Convention, False)
	FirstOfMonth_date_adjusted = ql_to_string(QlFirstOfMonth_date_adjusted)
	
	reportingDF[Underlying[0]].at['Open','MTD']  = 100*(time_series[Underlying[0]].at[End_date,'OPEN'] - time_series[Underlying[0]].at[FirstOfMonth_date_adjusted,'OPEN'])/time_series[Underlying[0]].at[FirstOfMonth_date_adjusted,'OPEN']
	reportingDF[Underlying[0]].at['Close','MTD'] = 100*(time_series[Underlying[0]].at[End_date,'CLOSE'] - time_series[Underlying[0]].at[FirstOfMonth_date_adjusted,'CLOSE'])/time_series[Underlying[0]].at[FirstOfMonth_date_adjusted,'CLOSE']
	reportingDF[Underlying[0]].at['High','MTD']  = 100*(time_series[Underlying[0]].at[End_date,'HIGH'] - time_series[Underlying[0]].at[FirstOfMonth_date_adjusted,'HIGH'])/time_series[Underlying[0]].at[FirstOfMonth_date_adjusted,'HIGH']
	reportingDF[Underlying[0]].at['Low','MTD']   = 100*(time_series[Underlying[0]].at[End_date,'LOW'] - time_series[Underlying[0]].at[FirstOfMonth_date_adjusted,'LOW'])/time_series[Underlying[0]].at[FirstOfMonth_date_adjusted,'LOW']
	reportingDF[Underlying[0]].at['VolK','MTD']  = 100*(time_series[Underlying[0]].at[End_date,'VOLK'] - time_series[Underlying[0]].at[FirstOfMonth_date_adjusted,'VOLK'])/time_series[Underlying[0]].at[FirstOfMonth_date_adjusted,'VOLK']
	reportingDF[Underlying[0]].at['VolB','MTD']  = 100*(time_series[Underlying[0]].at[End_date,'VOLB'] - time_series[Underlying[0]].at[FirstOfMonth_date_adjusted,'VOLB'])/time_series[Underlying[0]].at[FirstOfMonth_date_adjusted,'VOLB']
	
	#VAR YTD
	QlNewYear_date_adjusted = Calendar.advance(ql.Date(1,1,QlEnd_date.year()), Maturity, Convention, False)
	NewYear_date_adjusted = ql_to_string(QlNewYear_date_adjusted)
	
	reportingDF[Underlying[0]].at['Open','YTD']  = 100*(time_series[Underlying[0]].at[End_date,'OPEN'] - time_series[Underlying[0]].at[NewYear_date_adjusted,'OPEN'])/time_series[Underlying[0]].at[NewYear_date_adjusted,'OPEN']
	reportingDF[Underlying[0]].at['Close','YTD'] = 100*(time_series[Underlying[0]].at[End_date,'CLOSE'] - time_series[Underlying[0]].at[NewYear_date_adjusted,'CLOSE'])/time_series[Underlying[0]].at[NewYear_date_adjusted,'CLOSE']
	reportingDF[Underlying[0]].at['High','YTD']  = 100*(time_series[Underlying[0]].at[End_date,'HIGH'] - time_series[Underlying[0]].at[NewYear_date_adjusted,'HIGH'])/time_series[Underlying[0]].at[NewYear_date_adjusted,'HIGH']
	reportingDF[Underlying[0]].at['Low','YTD']   = 100*(time_series[Underlying[0]].at[End_date,'LOW'] - time_series[Underlying[0]].at[NewYear_date_adjusted,'LOW'])/time_series[Underlying[0]].at[NewYear_date_adjusted,'LOW']
	reportingDF[Underlying[0]].at['VolK','YTD']  = 100*(time_series[Underlying[0]].at[End_date,'VOLK'] - time_series[Underlying[0]].at[NewYear_date_adjusted,'VOLK'])/time_series[Underlying[0]].at[NewYear_date_adjusted,'VOLK']
	reportingDF[Underlying[0]].at['VolB','YTD']  = 100*(time_series[Underlying[0]].at[End_date,'VOLB'] - time_series[Underlying[0]].at[NewYear_date_adjusted,'VOLB'])/time_series[Underlying[0]].at[NewYear_date_adjusted,'VOLB']
	

#Display




	
	
	
 
	
