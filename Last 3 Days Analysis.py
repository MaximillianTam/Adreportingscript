import openpyxl 
import pandas as pd
import random

#Keep this file 
# df = pd.read_excel('C:/Users/mysim/OneDrive/Desktop/Samsung 2022 Jan Ads.xlsx')

# df = pd.read_excel('C:/Users/mysim/OneDrive/Desktop/My Simple Essential - The Anti Case-Ad Report(2022-05-27 to 2022-05-30).xlsx')

df = pd.read_excel('C:/Users/mysim/OneDrive/Desktop/My Simple Essential - The Anti Case-Ad Report(2022-05-31 to 2022-06-02) (1).xlsx')

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
total = cpa_Table_ads['Cost'].sum()
cpa_Table_ads['Cost % to TTL']  = round((cpa_Table_ads['Cost'])/(total) * 100)
print(cpa_Table_ads[['Cost','Cost % to TTL','CPA','Total Complete Payment']])

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
cpa_Table_adgroups = cpa_Table_adgroups.reset_index(level = 0)
cpa_Table_adgroups = cpa_Table_adgroups.reset_index(level = 0)
print(cpa_Table_adgroups)

output3 = pd.pivot_table(df,
    index = ['Ad Group Name','Ad Name'],
    values = ['Cost','Total Complete Payment'],
    aggfunc = 'sum',
    fill_value=0
    )
output3['CPA'] = round(output3['Cost']/output3['Total Complete Payment'],2)


ad_adgroupcombo = output3[(output3['Cost'] != 0) ]
# ad_adgroupcombo = ad_adgroupcombo.reset_index(level = 0)

# cpa_Table_ads.to_excel('C:/Users/mysim/OneDrive/Desktop/Manual Bid Ad Performance.xlsx', index = False, header=True)
# ad_adgroupcombo.to_excel('C:/Users/mysim/OneDrive/Desktop/Manual Bid Ad Performance.xlsx', index = False, header=True)


print(ad_adgroupcombo)
