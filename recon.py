import pandas as pd
output_workbook = 'output.xlsx'

#read excel file
dfPortal = pd.read_excel('Portal.xlsx')
dfStore = pd.read_excel('Store.xlsx')

#find duplicate Entries in Portal
dfPortalDup = dfPortal[dfPortal.duplicated('Bill No')]

#joing two sheets on bill no column
df1 = pd.merge(dfStore ,dfPortal[['Bill No','PortalGross','PortalTax','PortalNet']],
 on = 'Bill No',indicator = True, how='outer')

#extract bills that are present only in Store report
dfStoreOnly = df1[df1['_merge'] == 'left_only' ]
dfStoreOnly.drop(['_merge'],axis=1,inplace=True)
#extract bills that are present only in Store report
dfPortalOnly = df1[df1['_merge'] == 'right_only' ]
dfPortalOnly.drop(['_merge'],axis=1,inplace=True)

#extract bills that are present in both reports
dfCommons = df1[df1['_merge'] == 'both' ]
dfCommons.drop(['_merge'],axis=1,inplace=True)

#find bills with difference in Gross sale
#find difference in gross sale and then find absolute value of that difference 
dfCommons = dfCommons.assign(GrossDiff=(dfCommons['StoreGross'] - dfCommons['PortalGross']))
dfCommons = dfCommons.assign(GrossDiffAbs=(dfCommons['GrossDiff'].abs()))
dfCommons.drop(['GrossDiff'],axis=1,inplace=True) 
#filter table where gross difference in greater than 1, taking rounding in to consideration
dfGrossDiff = dfCommons[dfCommons['GrossDiffAbs'] > 1 ]
dfGrossDiff.loc['Total', 'GrossDiffAbs'] = dfGrossDiff.GrossDiffAbs.sum()


#find bills with difference in Net sale
#find difference in Net sale and then find absolute value of that difference 
dfCommons.drop(['GrossDiffAbs'],axis=1,inplace=True) 
dfCommons = dfCommons.assign(NetDiff=(dfCommons['StoreNet'] - dfCommons['PortalNet']))
dfCommons = dfCommons.assign(NetDiffAbs=(dfCommons['NetDiff'].abs()))
dfCommons.drop(['NetDiff'],axis=1,inplace=True) 
#filter table where Net  difference in greater than 1, taking rounding in to consideration
dfNetDiff = dfCommons[dfCommons['NetDiffAbs'] > 1 ]
dfNetDiff.loc['Total', 'NetDiffAbs'] = dfNetDiff.NetDiffAbs.sum()

#find bills with difference in Tax
#find difference in Tax and then find absolute value of that difference 
dfCommons.drop(['NetDiffAbs'],axis=1,inplace=True) 
dfCommons = dfCommons.assign(TaxDiff=(dfCommons['StoreTax'] - dfCommons['PortalTax']))
dfCommons = dfCommons.assign(TaxDiffAbs=(dfCommons['TaxDiff'].abs()))
dfCommons.drop(['TaxDiff'],axis=1,inplace=True) 
#filter table where tax   difference in greater than 1, taking rounding in to consideration
dfTaxDiff = dfCommons[dfCommons['TaxDiffAbs'] > 1 ]
dfTaxDiff.loc['Total', 'TaxDiffAbs'] = dfTaxDiff.TaxDiffAbs.sum()




with pd.ExcelWriter('output.xlsx') as writer:
    dfStoreOnly.to_excel(writer,sheet_name='StoreOnly',index=False)
    dfPortalOnly.to_excel(writer,sheet_name='PortalOnly',index=False)
    dfGrossDiff.to_excel(writer,sheet_name='GrossDifference',index=False)
    dfNetDiff.to_excel(writer,sheet_name='NetDifference',index=False)
    dfTaxDiff.to_excel(writer,sheet_name='TaxDifference',index=False)
    dfPortalDup.to_excel(writer,sheet_name='DuplicatesInPortal',index=False)


