# Import modules
import xlwings as xw
import datetime as dt
import win32api
import matplotlib.pyplot as plt
import pandas as pd

def main():
    wb = xw.Book.caller()
    # wb.sheets[0].range("A1").value = "Hello xlwings!"		# test code

    #--------------------------------------------------------------------------------------------------------------------------------
    # Define sheets
    sht_asfe1_cp = wb.sheets['ASFE1-CP']
    sht_asfe1_er = wb.sheets['ASFE1-ER']
    sht_cp_plot = wb.sheets['CP Plot']
    sht_er_plot = wb.sheets['ER Plot']

    #--------------------------------------------------------------------------------------------------------------------------------
    # Draw CP Plot
    fig_cp = plt.figure()
    df_cp = sht_asfe1_cp.range('A10').options(
        pd.DataFrame, header=1, index=False, expand='table'
        ).value											# fetch the data from sheet- 'ASFE1-CP'
    # sht_cp_plot.range('A1').options(index=False).value = df_cp   	# show the dataframe values into sheet- 'CP Plot'
    plt.plot(df_cp["Date (MM/DD/YY)"], df_cp["delta CP"])
    sht_cp_plot.pictures.add(fig_cp, name= "ASFE1_CP_Plot", update= True)

    #--------------------------------------------------------------------------------------------------------------------------------
    # Draw ER Plot
    # fig_er = plt.figure()
    # df_er = sht_asfe1_er.range('A9').options(
        # pd.read_excel("H:\\excel\\dryetch\\macro_enabled_logbooks\\ASH09_QC_LOG_BOOK\\ASH09_QC_LOG_BOOK.xlsm", sheetname='ASFE1-ER'), header=1, index=False
        # ).value
    df_er = sht_asfe1_er.range('A9').options(
        pd.DataFrame, header=1, index=False
        ).value
            
    # df_er.replace('', np.NaN, inplace=True)        # replace the empty boxes with spaces
    sht_er_plot.range('A1').options(index=False).value = df_er       # show the dataframe values into sheet- 'CP Plot'
    # sht_er_plot.range('A1').options(index=False).value = sht_asfe1_er.range('A9').value
    # plt.plot(df_er["Date (MM/DD/YY)"], df_er["Etch Rate (A/Min)"])
    # sht_er_plot.pictures.add(fig_er, name= "ASFE1_ER_Plot", update= True)




#--------------------------------------------------------------------------------------------------------------------------------
# User Defined Functions (UDFs)
#--------------------------------------------------------------------------------------------------------------------------------
@xw.func
def hello(name):
    return "hello {0}".format(name)



## REFERENCES
# https://stackoverflow.com/questions/22937650/pandas-reading-excel-with-merged-cells