import os
import glob
import pandas as pd
import calendar

from datetime import datetime

# get file path
file_path = os.getcwd()
otc_invoice_files = glob.glob(os.path.join(file_path, "*.xls"))

# create list
trans_date = []
total_collected = []
total_due = []

# read xls file one by one
for invoice_file in otc_invoice_files:
    # load xls file to DataFrame
    df = pd.read_excel(invoice_file)

    # get transaction date and time
    # since the last transaction row in the last but two row, I use '-3' to get the cell value
    last_trans_date = df['Trans Date'].iloc[-3]
    last_trans_time = df['Trans Time'].iloc[-3]

    invoice_date_time = datetime.strptime(
        last_trans_date+" "+last_trans_time, "%m-%d-%Y %H:%M:%S")
    invoice_date = invoice_date_time.strftime('%Y-%m-%d')

    # get amount from last row
    amount_collected = df['Total Collected'].iloc[-1]
    amount_due = df['Total Due'].iloc[-1]

    # add value to list
    trans_date.append(invoice_date)
    total_collected.append(amount_collected)
    # due amount is neg, so convert to postive
    postive_due_amount = abs(amount_due)
    total_due.append(postive_due_amount)

# create a new DataFrame
# create a new data frame
column_name = ['Date', 'Collected', 'Due']
report_df = pd.DataFrame(columns=column_name)

# add list to DataFrame
report_df["Date"] = trans_date
report_df["Collected"] = total_collected
report_df["Due"] = total_due

# sort by Date ASC, and update daframe force
report_df.sort_values(by='Date', inplace=True)

# get transaction date from last node
last_row_date = trans_date[len(trans_date)-1]

# get year and month
last_row_date_year = int(last_row_date[:4])
last_row_date_month = int(last_row_date[5:7])

# get days in this month
month_days = calendar.monthrange(
    int(last_row_date_year), int(last_row_date_month))[1]


month_report_df = pd.DataFrame(columns=column_name)

#month_report_df = pd.DataFrame()

for i in range(1, month_days+1):
    selected_date = last_row_date[:4]+'-' + \
        last_row_date[5:7]+'-'+str(i).zfill(2)

    temp_df=report_df.query("Date=='"+selected_date+"'")
    
    
    if(temp_df.empty):
        month_report_df.loc[len(month_report_df)] = [selected_date, '0', '0']
    else:
        temp_date=temp_df.iloc[0]["Date"]
        temp_collect=str(temp_df.iloc[0]["Collected"])
        temp_due=str(temp_df.iloc[0]["Due"])
        #print(temp_due)
        month_report_df.loc[len(month_report_df)] = [temp_date,temp_collect,temp_due]

    #print(str(i)+'--'+"Size = "+str(temp_df.empty))

    #if(len(df_new) == 1):
        #month_report_df.loc[len(month_report_df)] = df_new
    #else:
        #month_report_df.loc[len(month_report_df)] = [selected_date, '0', '0']
print('\n\n')
print(month_report_df)


# generate csv file, and remove index column
csv_name = 'OTC_report.csv'
#report_df.to_csv(csv_name, index=False)
