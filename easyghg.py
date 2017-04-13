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
    main = xw.sheets[0]
    full_output = xw.sheets[1]
    full_output.clear_contents()

    years = main.range('A3').options(np.array, expand='vertical').value


    co2_emiss = np.zeros_like(years)
    ch4_emiss = np.zeros_like(years)

    co2 = main.range('A3').offset(column_offset=1).options(np.array, empty=0, expand='vertical').value
    ch4 = main.range('A3').offset(column_offset=2).options(np.array, empty=0, expand='vertical').value

    co2_emiss[:len(co2)] = co2
    ch4_emiss[:len(ch4)] = ch4


    ch4_fossil = str_to_bool(main.range('N2').value)
    ch4_fb = str_to_bool(main.range('N3').value)
    pulse = str_to_bool(main.range('N11').value)
    n_runs = int(main.range('N6').value)
    RS = int(main.range('N7').value)
    full_out = str_to_bool(main.range('N8').value)
    kind = main.range('N12').value

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


        full_output.range('A2').value = total_df
        full_output.range('A2').value = 'Year'
        full_output.range('B1').value = 'Runs'

        mean_rf = co2_mc['mean'] + ch4_mc['mean']
        up_sigma = co2_mc['+sigma'] + ch4_mc['+sigma']
        down_sigma = co2_mc['-sigma'] + ch4_mc['-sigma']

        main.range('E3').options(transpose=True).value = co2_mc['mean'].values
        main.range('F3').options(transpose=True).value = ch4_mc['mean'].values
        main.range('G3').options(transpose=True).value = mean_rf.values

        sns.set_style('whitegrid')
        plt.plot(mean_rf)
        plt.fill_between(years, up_sigma, down_sigma, alpha=0.3)
        plt.xlabel('Years', size=14)

        if kind == 'RF':
            plt.ylabel('$W \ m^{-2}$', size=14)
            plt.title('Mean Forcing\nwith 1-sigma uncertainty', size=15)
        elif kind == 'CRF':
            plt.ylabel('$W \ m^{-2} \ y$', size=14)
            plt.title('Mean Cumulative Forcing\nwith 1-sigma uncertainty', size=15)
        sns.despine()
        fig = plt.gcf()
        # plot = main.Plot(fig).show('MyPlot',
        #                         left=xw.range('I15').left,
        #                         top=xw.range('I15').top,
        #                         width=533, height=400)

        main.pictures.add(fig, name='MyPlot', update=True,
                     left=main.range('I15').left, top=main.range('I15').top,
                     width=533, height=400)

    else:
        co2_rf = CO2(co2_emiss, years,  kind=kind,)
        ch4_rf = CH4(ch4_emiss, years,  kind=kind,
                    decay=ch4_fossil, cc_fb=ch4_fb)
        total = co2_rf + ch4_rf

        main.range('E3').options(transpose=True).value = co2_rf
        main.range('F3').options(transpose=True).value = ch4_rf
        main.range('G3').options(transpose=True).value = total
        main.range('U3').options(transpose=True).value = ch4_emiss


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
        # plot = main.Plot(fig).show('MyPlot',
        #                         left=xw.range('I15').left,
        #                         top=xw.range('I15').top,
        #                         width=533, height=400)
        main.pictures.add(fig, name='MyPlot', update=True,
                     left=main.range('I15').left, top=main.range('I15').top,
                     width=533, height=400)
