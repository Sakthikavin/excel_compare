import pandas as pd
import numpy as np

# Define the diff function to show the changes in each field
def report_diff(x):
    return x[0] if x[0] == x[1] else '{} ---> {}'.format(*x)

def get_combined_index(df, key_col_list):
    df['combined_index'] = ''
    for col in key_col_list:
        df['combined_index'] = df['combined_index'].astype(str) + '_' + df[col].astype(str)
    return df


columns = ['model_name', 'extra','Total Tests Ran','Total Tests Passed','Total Tests Failed','Test Fail Percentage']
key_columns = ['model_name', 'extra']

# Read in the two files but call the data old and new and create columns to track
old = pd.read_excel('A.xlsx', 'Evaluation Summary', na_values=['NA'])
new = pd.read_excel('B.xlsx', 'Evaluation Summary', na_values=['NA'])
old['status'] = 'old'
new['status'] = 'new'

old = get_combined_index(old, key_columns)
new = get_combined_index(new, key_columns)

#Join all the data together and ignore indexes so it all gets added
full_set = pd.concat([old,new],ignore_index=True)

# Let's see what changes in the main columns we care about
changes = full_set.drop_duplicates(subset=columns,keep='last')

#We want to know where the duplicate row numbers are, that means there have been changes
dupe_accts = changes.set_index('combined_index').index.get_duplicates()

#Get all the duplicate rows
dupes = changes[changes["combined_index"].isin(dupe_accts)]

#Pull out the old and new data into separate dataframes
change_new = dupes[(dupes['status'] == 'new')]
change_old = dupes[(dupes['status'] == 'old')]

#Drop the temp columns - we don't need them now
change_new = change_new.drop(['status'], axis=1)
change_old = change_old.drop(['status'], axis=1)

#Index on the unique set of columns
change_new.set_index('combined_index',inplace=True, drop=False)
change_old.set_index('combined_index',inplace=True, drop=False)

#Now we can diff because we have two data sets of the same size with the same index
diff_panel = pd.Panel(dict(df1=change_old,df2=change_new))
modified_rows = diff_panel.apply(report_diff, axis=0)
modified_rows['status'] = 'modified' 

#Diff'ing is done, we need to get a list of removed items

#Flag all duplicated row numbers
changes['duplicate']=changes["combined_index"].isin(dupe_accts)

#Identify non-duplicated items that are in the old status and did not show in the new status
removed_rows = changes[(changes['duplicate'] == False) & (changes['status'] == 'old')]
removed_rows = removed_rows.drop(['duplicate'], axis=1)

# We have the old and diff, we need to figure out which ones are new

#Drop duplicates but keep the first item instead of the last
new_row_set = full_set.drop_duplicates(subset=columns,keep='first')

#Identify dupes in this new dataframe
new_row_set['duplicate']=new_row_set["combined_index"].isin(dupe_accts)

#Identify added rows
added_rows = new_row_set[(new_row_set['duplicate'] == False) & (new_row_set['status'] == 'new')]
added_rows = added_rows.drop(['duplicate'], axis=1)

# Concat all added, removed and modified
consolidated = pd.concat([modified_rows, removed_rows, added_rows], ignore_index=True, sort=False)


# Get the unchanged rows from the current dataframe
unchanged_rows = new[~new.combined_index.isin(consolidated.combined_index)]

#Save the changes to excel but only include the columns we care about
writer = pd.ExcelWriter('get_combined_index.xlsx')

modified_rows.to_excel(writer,'modified',index=False)
removed_rows.to_excel(writer,'removed',index=False)
added_rows.to_excel(writer,'added',index=False)
consolidated.to_excel(writer, 'consolidated', index=False)
writer.save()