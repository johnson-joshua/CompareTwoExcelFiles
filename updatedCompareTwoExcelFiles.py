# -*- coding: utf-8 -*-
"""
Created on Tue Dec 11 09:14:50 2018

@author: jmjohnson-zonetech
"""

import pandas as pd
from pathlib import Path


def excel_differences(path_OLD, path_NEW):
    xlOld = pd.ExcelFile(path_OLD)
    xlNew = pd.ExcelFile(path_NEW)
    dfOld = xlOld.parse('Sheet1')
    dfNew = xlNew.parse('Sheet1')
    writer = pd.ExcelWriter('Differences.xlsx', engine='xlsxwriter')
    sheetNameOld = path_OLD.stem
    sheetNameNew = path_NEW.stem
    
    myOldList = dfOld['sis_id']
    myNewList = dfNew['sis_id']
    
#Find differences
    dfNotInOld = dfNew.loc[~dfNew['sis_id'].isin(myOldList)]
#    print(dfNotInOld.iloc[:,0])
    dfNotInNew = dfOld.loc[~dfOld['sis_id'].isin(myNewList)]
    dfNotInOld.to_excel(writer, sheet_name='Not In ' +sheetNameOld, index=False)
    dfNotInNew.to_excel(writer, sheet_name='Not In ' +sheetNameNew, index=False)
    remove_differences(dfOld, dfNotInOld, dfNew, dfNotInNew)
    writer.save()

def remove_differences(dfOld, dfNotInOld, dfNew, dfNotInNew):
    #Removes the rows that aren't in the other dataframe
    dfOld = dfOld.drop(dfNotInNew.index, axis=0)
    dfNew = dfNew.drop(dfNotInOld.index, axis=0)
    excel_changes(dfOld, dfNew)
    
def excel_changes(dfOld, dfNew):
    #Sort both dataframes by ID
    dfOld = dfOld.sort_values(by='sis_id',  ascending=True)
    dfNew = dfNew.sort_values(by='sis_id',  ascending=True)
    # Perform Diff
    dfDiff = dfOld.copy()
    for row in range(dfDiff.shape[0]):
        for col in range(dfDiff.shape[1]):
            value_OLD = dfOld.iloc[row,col]
            value_NEW = dfNew.iloc[row,col]
            if value_OLD==value_NEW:
                dfDiff.iloc[row,col] = dfNew.iloc[row,col]
            else:
                dfDiff.iloc[row,col] = ('{}→{}').format(value_OLD,value_NEW)
    # Save output and format
    writer = pd.ExcelWriter('Changes.xlsx', engine='xlsxwriter')

    dfDiff.to_excel(writer, sheet_name='ValuesThatHaveChanged', index=False)
    
    workbook  = writer.book
    worksheet = writer.sheets['ValuesThatHaveChanged']
    worksheet.hide_gridlines(2)
    
    grey_fmt = workbook.add_format({'font_color': '#E0E0E0'})
    highlight_fmt = workbook.add_format({'font_color': '#FF0000', 'bg_color':'#B1B3B3'})
    
                                             ## highlight changed cells
    worksheet.conditional_format('A1:ZZ10000', {'type': 'text',
                                            'criteria': 'containing',
                                            'value':'→',
                                            'format': highlight_fmt})
    ## highlight unchanged cells
    worksheet.conditional_format('A1:ZZ10000', {'type': 'text',
                                            'criteria': 'not containing',
                                            'value':'→',
                                            'format': grey_fmt})
    # save
    writer.save()
    print('Done with changes.')
    
     
def main():
    path_OLD = Path(r'C:\Users\jmjohnson-zonetech\Desktop\MyOnStudentsSQL.xlsx')
    path_NEW = Path(r'C:\Users\jmjohnson-zonetech\Desktop\MyOnStudentsPQ.xlsx')

    excel_differences(path_OLD, path_NEW)
    print('Done with differences.')

    
if __name__ == '__main__':
    main()