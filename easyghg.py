from __future__ import division
import xlwings as xw
import pandas as pd
import numpy as np
from ghgforcing import CO2, CH4
import matplotlib.pyplot as plt
import seaborn as sns

def str_to_bool(s):
    if s == 'True':
         return True
    elif s == 'False':
         return False
    else:
         raise ValueError # evil ValueError that doesn't tell you what the wrong value was

def ghg_calc():
    wb = xw.Workbook.caller()
    xw.Sheet('Full_output').clear_contents()
    tstep = 1
    years = xw.Range('A3').vertical.options(np.array).value

    co2_emiss = xw.Range('A3').vertical.offset(column_offset=1).options(np.array,
                                                                        empty=0).value

    ch4_emiss = xw.Range('A3').vertical.offset(column_offset=2).options(np.array,
                                                                        empty=0).value

    ch4_fossil = str_to_bool(xw.Range('N2').value)
    ch4_fb = str_to_bool(xw.Range('N3').value)
    pulse = str_to_bool(xw.Range('N11').value)
    n_runs = int(xw.Range('N6').value)
    RS = int(xw.Range('N7').value)
    full_out = str_to_bool(xw.Range('N8').value)
    kind = xw.Range('N12').value

    if pulse == True:
        tstep = 0.0001 #very small tstep to approximate a delta function
        co2_init = co2_emiss[0]
        ch4_init = ch4_emiss[0]
        
        years = np.linspace(0,100, 100/tstep+1)
        co2_emiss = np.zeros_like(years)
        co2_emiss[0] = co2_init
        
        ch4_emiss = np.zeros_like(years)
        ch4_emiss[0] = ch4_init
    
    if n_runs > 1:
        
        co2_mc, co2_full = CO2(co2_emiss, years, RS=RS, kind=kind,
                    runs=n_runs, full_output=full_out)
        ch4_mc, ch4_full = CH4(ch4_emiss, years, RS=RS, kind=kind, 
                    runs=n_runs, full_output=full_out,
                    decay=ch4_fossil, cc_fb=ch4_fb)
        total_full = co2_full + ch4_full
        total_df = pd.DataFrame(total_full)
        
        
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
        plt.fill_between(years, up_sigma, down_sigma, alpha=0.4)
        plt.xlabel('Years', size=14)
        
        if kind == 'RF':
            plt.ylabel('$W \ m^{-2}$', size=14)
            plt.title('Mean Forcing\nwith 1-sigma uncertainty', size=15)
        elif kind == 'CRF':
            plt.ylabel('$W \ m^{-2} \ y$', size=14)
            plt.title('Mean Cumulative Forcing\nwith 1-sigma uncertainty', size=15)
        sns.despine()
        fig = plt.gcf()
        plot = xw.Plot(fig).show('MyPlot', 
                                left=xw.Range('I15').left, 
                                top=xw.Range('I15').top,
                                width=533, height=400)
    
    else:
        co2_rf = CO2(co2_emiss, years,  kind=kind,)
        ch4_rf = CH4(ch4_emiss, years,  kind=kind,
                    decay=ch4_fossil, cc_fb=ch4_fb)
        total = co2_rf + ch4_rf
        
        xw.Range('E3').options(transpose=True).value = co2_rf
        xw.Range('F3').options(transpose=True).value = ch4_rf
        xw.Range('G3').options(transpose=True).value = total
        xw.Range('U3').options(transpose=True).value = ch4_emiss
        
        
        sns.set_style('whitegrid')
        plt.plot(total)
        plt.xlabel('Years', size=14)
        if kind == 'RF':
            plt.ylabel('$W \ m^{-2}$', size=14)
            plt.title('Forcing', size=15)
        elif kind == 'CRF':
            plt.ylabel('$W \ m^{-2} \ y$', size=14)
            plt.title('Cumulative Forcing', size=15)
        sns.despine()
        fig = plt.gcf()
        plot = xw.Plot(fig).show('MyPlot', 
                                left=xw.Range('I15').left, 
                                top=xw.Range('I15').top,
                                width=533, height=400)
