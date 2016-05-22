from __future__ import division
import xlwings as xw
import pandas as pd
import numpy as np
from ghgforcing import CO2, CH4
import matplotlib.pyplot as plt
import seaborn as sns


def ghg_calc():
    wb = xw.Workbook.caller()
    years = np.arange(0,101)
    co2_emiss = xw.Range('B3:B103').options(np.array,empty=0).value
    ch4_emiss = xw.Range('C3:C103').options(np.array,empty=0).value
    ch4_fossil = xw.Range('N2').value
    ch4_fb = xw.Range('N3').value
    n_runs = int(xw.Range('N6').value)
    RS = int(xw.Range('N7').value)
    full_out = xw.Range('N8').value
    
    #if len(co2_emiss) < len(years):
    #    diff_co2 = len(years) - len(co2_emiss)
    #    co2_emiss.append(np.zeros(diff_co2))
    
    #if len(ch4_emiss) < len(years):
    #    diff_ch4 = len(years) - len(ch4_emiss)
    #    ch4_emiss.append(np.zeros(diff_ch4))
    
    
    xw.Range('B3').options(transpose=True).value = co2_emiss
    xw.Range('C3').options(transpose=True).value = ch4_emiss
    
    if n_runs > 1:
        
        co2_mc, co2_full = CO2(co2_emiss, years, tstep=1, RS=RS, 
                    runs=n_runs, full_output=full_out)
        ch4_mc, ch4_full = CH4(ch4_emiss, years, tstep=1, RS=RS, 
                    runs=n_runs, full_output=True)
        total_full = co2_full + ch4_full
        total_df = pd.DataFrame(total_full)
        
        #wb = xw.Workbook('Full_output')
        
        xw.Range('Full_output', 'A2').value = total_df
        xw.Range('Full_output', 'A2').value = 'Year'
        xw.Range('Full_output', 'B1').value = 'Runs'
        
        mean_rf = co2_mc['mean'] + ch4_mc['mean']
        up_sigma = co2_mc['+sigma'] + ch4_mc['+sigma']
        down_sigma = co2_mc['-sigma'] + ch4_mc['-sigma']
        
        xw.Range('E3').options(transpose=True).value = co2_mc['mean'].values
        xw.Range('F3').options(transpose=True).value = ch4_mc['mean'].values
        xw.Range('G3').options(transpose=True).value = mean_rf.values
        
        sns.set_style('whitegrid')
        plt.plot(mean_rf)
        plt.fill_between(years, up_sigma, down_sigma, alpha=0.5)
        plt.xlabel('Years', size=14)
        plt.ylabel('$W \ m^{-2}$', size=14)
        sns.despine()
        fig = plt.gcf()
        plot = xw.Plot(fig).show('MyPlot', 
                                left=xw.Range('I13').left, 
                                top=xw.Range('I13').top,
                                width=533, height=400)
    
    else:
        co2_rf = CO2(co2_emiss, years, tstep=1)
        ch4_rf = CH4(ch4_emiss, years, tstep=1,
                    decay=ch4_fossil, cc_fb=ch4_fb)
        total = co2_rf + ch4_rf
        
        xw.Range('E3').options(transpose=True).value = co2_rf
        xw.Range('F3').options(transpose=True).value = ch4_rf
        xw.Range('G3').options(transpose=True).value = total
        
        
        sns.set_style('whitegrid')
        plt.plot(total)
        plt.xlabel('Years', size=14)
        plt.ylabel('$W \ m^{-2}$', size=14)
        sns.despine()
        fig = plt.gcf()
        plot = xw.Plot(fig).show('MyPlot', 
                                left=xw.Range('I13').left, 
                                top=xw.Range('I13').top,
                                width=533, height=400)
    
