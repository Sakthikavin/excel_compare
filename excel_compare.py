import pandas as pd
import numpy as np

# Define the diff function to show the changes in each field
def report_diff(x):
    return x[0] if x[0] == x[1] else '{} ---> {}'.format(*x)

columns = ['model_name','Total Tests Ran','Total Tests Passed','Total Tests Failed','Test Fail Percentage']
key_column = ['model_name']

# Read in the two files but call the data old and new and create columns to track
old = pd.read_excel('A.xlsx', 'Evaluation Summary', na_values=['NA'])
new = pd.read_excel('B.xlsx', 'Evaluation Summary', na_values=['NA'])
old['status'] = 'old'
new['status'] = 'new'

#Join all the data together and ignore indexes so it all gets added
full_set = pd.concat([old,new],ignore_index=True)

# Let's see what changes in the main columns we care about
changes = full_set.drop_duplicates(subset=columns,keep='last')

#We want to know where the duplicate account numbers are, that means there have been changes
dupe_accts = changes.set_index('model_name').index.get_duplicates()

#Get all the duplicate rows
dupes = changes[changes["model_name"].isin(dupe_accts)]

#Pull out the old and new data into separate dataframes
change_new = dupes[(dupes['status'] == 'new')]
change_old = dupes[(dupes['status'] == 'old')]

#Drop the temp columns - we don't need them now
change_new = change_new.drop(['status'], axis=1)
change_old = change_old.drop(['status'], axis=1)

#Index on the unique set of columns
change_new.set_index('model_name',inplace=True, drop=False)
change_old.set_index('model_name',inplace=True, drop=False)

#Now we can diff because we have two data sets of the same size with the same index
diff_panel = pd.Panel(dict(df1=change_old,df2=change_new))
diff_output = diff_panel.apply(report_diff, axis=0)
diff_output['status'] = 'duplicate' 

#Diff'ing is done, we need to get a list of removed items

#Flag all duplicated account numbers
changes['duplicate']=changes["model_name"].isin(dupe_accts)

#Identify non-duplicated items that are in the old status and did not show in the new status
removed_accounts = changes[(changes['duplicate'] == False) & (changes['status'] == 'old')]
removed_accounts = removed_accounts.drop(['duplicate'], axis=1)

# We have the old and diff, we need to figure out which ones are new

#Drop duplicates but keep the first item instead of the last
new_account_set = full_set.drop_duplicates(subset=columns,keep='first')

#Identify dupes in this new dataframe
new_account_set['duplicate']=new_account_set["model_name"].isin(dupe_accts)

#Identify added accounts
added_accounts = new_account_set[(new_account_set['duplicate'] == False) & (new_account_set['status'] == 'new')]
added_accounts = added_accounts.drop(['duplicate'], axis=1)

# Concat all added, removed and modified
consolidated = pd.concat([diff_output, removed_accounts, added_accounts], ignore_index=True, sort=False)

#Save the changes to excel but only include the columns we care about
writer = pd.ExcelWriter('my-diff-2.xlsx')

diff_output.to_excel(writer,'changed',index=False)
removed_accounts.to_excel(writer,'removed',index=False)
added_accounts.to_excel(writer,'added',index=False)
consolidated.to_excel(writer, 'consolidated', index=False)
writer.save()