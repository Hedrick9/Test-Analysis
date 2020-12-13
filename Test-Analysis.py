## ~~ IMPORT PACKAGES ~~
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
## ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# Input Parameters
TC_list = [11, 12, 13, 14] # list of thermocouple channels used
Final_Temp = 200

## ~~ REMOVE UNUSED HEADERS ~~
#headers = []
#for item in df.columns:
#    if 'Temp' in item:
#        headers.append(item)
#    else:
#        pass
#for i in range(8, len(headers)):
#   df = df.drop(headers[i], axis=1)

## ~~ FUNCTIONS FOR ANALYSIS ~~
def Wok_Results(df):
    # Determine Test Start TIme
    tst_i = 0 # start time index
    for i in range(len(df)):
        if (df.iloc[i+1, TC_list].mean() - df.iloc[i, TC_list].mean()) > 0.5:
            tst_i = i
            break
        else:
            pass

    tst = df['Test Time'][tst_i] # test start time
    # Determine Test End Time
    tet_i = 0 # test end time index
    for i in range(len(df)):
        if df.iloc[i, TC_list].mean() >= Final_Temp and df.iloc[i, TC_list].mean() < (Final_Temp + 10):
            tet_i = i
            break
        else:
            pass
    tet = df['Test Time'][tet_i] # test end time
    
    # ~~RESULTS~~
    t_test = tet - tst # total test time (min) 
    # Initial Temp
    Ti = df.iloc[tst_i, TC_list].mean() 
    # Final Temp
    Tf = df.iloc[tet_i, TC_list].mean() 
    # Energy Consumption
    E_ec = df['Wh'][tst_i:tet_i].sum() # Electric Energy Consumption (Wh)
    E_er = E_ec / (t_test/60) # Electric Energy Rate (W)
    # Energy to Product
    m_pan = 6.783 # mass of pan (lb)
    c_pan = 0.11 / 3.41214148 # specific heat capacity of pan (Wh/lbm-°F)
    m_water = 10 # mass of water (lb)
    c_water = 1 / 3.41214148 # specific heat capacity of water (Wh/lbm-°F)
    E_water = round(m_water*c_water*(Tf - Ti), 2)  # Energy to Water (Wh)
    E_pan = round(m_pan*c_pan*(Tf - Ti), 2)  # Energy to Pan (Wh)
    Eff = round(((E_water + E_pan)/E_ec)*100, 2) # Efficiency (%)
    # Extra
    V = round(df['Voltage'][tst_i:tet_i].mean(), 2) # Voltage 
    T_amb = round(df['Ambient'][tst_i:tet_i].mean(), 2) # Average Ambient Temperature
    
    return t_test, V, Ti, Tf, T_amb, m_pan, m_water, E_er, E_ec, E_pan, E_water, Eff

## ~~ANALYSIS ON TEST OF INTEREST ~~
# Test variables 
test_variables = ['Test Time', 'Voltage', 'Initial Temp', 'Final Temp', 'Average Ambient', 'Pan Weight', 'Water Weight', 'Energy Rate', 'Energy Consumption', 'Pan Energy', 'Water Energy', 'Efficiency']
test_var_units = ['min', 'V', '°F', '°F', '°F', 'lbs', 'lbs', 'W', 'Wh', 'Wh', 'Wh', '%']
# Dictionary conataining summary of results from each test
n_tests = 1 # number of tests
results_dict = {}
for i in range(1, n_tests+1):
    df = pd.read_excel('/Users/Russell/Desktop/12-09-20_pantin_{}.xlsx'.format(i), skiprows=(0, 1, 2, 3, 5), engine='openpyxl')
    results_dict['Test {}'.format(i)] = Wok_Results(df)

# Dataframe contatining formatted summary of results for each test
results_df = pd.DataFrame(np.array([test_variables,
                                    test_var_units]).T, columns = ['', 'units'])
index = 1 # starting column for results_df
for item in results_dict:
    results_df.insert(index, '{}'.format(item), results_dict[item])
    index += 1

## ~~ EXPORT RESULTS TO EXCEL FILE ~~
def to_Excel(file_name, results_df):

    xlwriter = pd.ExcelWriter('{}.xlsx'.format(file_name))
    results_df.to_excel(xlwriter, sheet_name='Results', index=False)
    xlwriter.close()

#to_Excel('Results', results_df)
print(results_df.head())
