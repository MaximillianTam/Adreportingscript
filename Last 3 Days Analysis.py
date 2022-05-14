import openpyxl 
import pandas as pd

# df = pd.read_excel('C:/Users/mysim/OneDrive/Desktop/My Simple Essential - The Anti Case-Ad Report(2022-05-03 to 2022-05-06).xlsx')

df = pd.read_excel('C:/Users/mysim/OneDrive/Desktop/My Simple Essential - The Anti Case-Ad Report(2022-05-11 to 2022-05-14) (1).xlsx')

df[['Ad Name', 'B']] = df['Ad Name'].str.split('-', 1, expand=True)

df[['Ad Group Name', 'C']] = df['Ad Group Name'].str.split('-', 1, expand=True)
# print(df)

#Pivot Table for Ads
output = pd.pivot_table(df,
    index = 'Ad Name',
    # columns = 'CATEGORY',
    values = ['Cost','Total Complete Payment'],
    aggfunc = 'sum',
    fill_value=0
    )

#adding a CPA column 
output['CPA'] = round(output['Cost']/output['Total Complete Payment'],2)
cpa_Table_ads = output[(output['Cost'] != 0)]
cpa_Table_ads.drop(cpa_Table_ads.tail(1).index,inplace=True)

cpa_Table_ads['CPA'].sort_values(ascending=True)

print(cpa_Table_ads) 


#Pivot Table for Ad Groups
output2 = pd.pivot_table(df,
    index = 'Ad Group Name',
    # columns = 'CATEGORY',
    values = ['Cost','Total Complete Payment'],
    aggfunc = 'sum',
    fill_value=0
    )

#adding a CPA column 
output2['CPA'] = round(output2['Cost']/output2['Total Complete Payment'],2)
cpa_Table_adgroups = output2[(output2['Cost'] != 0)]
cpa_Table_adgroups.drop(cpa_Table_adgroups.tail(1).index,inplace=True)
print(cpa_Table_adgroups)

output3 = pd.pivot_table(df,
    index = ['Ad Group Name','Ad Name'],
    # columns = 'CATEGORY',
    values = ['Cost','Total Complete Payment'],
    aggfunc = 'sum',
    fill_value=0
    )
output3['CPA'] = round(output3['Cost']/output3['Total Complete Payment'],2)
# print(output3.columns)
# output3 = output3.reset_index(level=0)
# print(output3.columns)

ad_adgroupcombo = output3[(output3['Cost'] != 0) ]
# output3[output3['CPA'].sort_values(ascending=True)]

# print(ad_adgroupcombo)
ad_adgroupcombo.to_excel('C:/Users/mysim/OneDrive - LVMH Fashion Group/Reporting/MISC/Ad Performance 05.10.xlsx', index = False, header=True)

# ad_adgroupcombo = ad_adgroupcombo.reset_index(level=0)

print(ad_adgroupcombo)