import subprocess
import sys

def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

#pip install jpype1 tabula-py numpy pandas matplotlib XlsxWriter
install('jpype1')
install('tabula-py')
install('numpy')
install('pandas')
install('matplotlib')
install('XlsxWriter')


# importing required modules
import os
import pandas as pd 
import datetime
import tabula
import numpy as np
import matplotlib.pyplot as plt


#pd.reset_option('all')

current_directory = str(os.getcwd())
file_list = os.listdir(current_directory)

pslips = []
failed_to_parse = []
months_processed = 0 
emp_id = input('Please enter emp id : ')
years = []

for file_name in file_list:
    oldfilepath = current_directory + '/' + file_name
    name, ext = os.path.splitext(file_name)
    if ext == '.pdf' and  file_name.startswith(emp_id):
        try: 
            year_var = name[-4:]
            years.append(year_var)
            
        except Exception as e:
            print('\nRan into error while parsing pdf names\n')
            print(str(e))
            continue


years = list(set(years))
years.sort()
months = ['JAN' , 'FEB' , 'MAR' , 'APR' , 'MAY' , 'JUN' , 'JUL' , 'AUG' , 'SEP' , 'OCT' , 'NOV' , 'DEC']


for year in years:
    for month in months: 
        try: 
            df = None
            df1 = None
            df2 = None
            df3 = None
            df4 = None



            pdf_path = current_directory + '/' + emp_id + '_' + month + '_' + year + '.pdf'
            file_name = emp_id + '_' + month + '_' + year + '.pdf'
            
            print('\nScraping file :  ' + pdf_path )
            table = tabula.read_pdf(file_name,pages=1)

            df = table[1]
            df.rename( columns={'Unnamed: 0':'Col1' , 'Earnings':'Col2' , 'No of Units':'Col3' , 'Earned':'Col4' , 'Deductions':'Col5' , 'Amount':'Col6' , 'Unnamed: 1':'Col7'  }, inplace=True )

            df1 = df.loc[:, ('Col1' , 'Col4')]
            df1.rename( columns={'Col1':'Category' , 'Col4':'Amount'}, inplace=True )
            df1 = df1.assign(Type = 'Addition')

            df2 = df.loc[:, ('Col5' , 'Col7')]
            df2.rename( columns={'Col5':'Category' , 'Col7':'Amount'}, inplace=True )
            df2 = df2.assign(Type = 'Deduction')
            
            try: 
                df3 = table[2].loc[:, ('Employer Contribution' , 'Earned')]
                df3.rename( columns={'Employer Contribution':'Category' , 'Earned':'Amount'}, inplace=True )
                df3 = df3.assign(Type = 'Deduction')
            except:
                pass
            
            try: 
                df4 = table[2].loc[:, ('Employer Contribution' , 'Earned')]
                df4.rename( columns={'Employer Contribution':'Category' , 'Earned':'Amount'}, inplace=True )
                df4 = df4.assign(Type = 'Addition')
            except:
                pass
            
          
            final_df = pd.concat([df1,df2,df3,df4])
            final_df['Amount'] = final_df['Amount'].fillna(0)
            final_df_final = final_df.loc[(final_df['Category'].notnull())]
            final_df_final = final_df_final.assign(Month_name  = month , Year_name = year )

            pslips.append(final_df_final)
            
            df = None
            df1 = None
            df2 = None
            df3 = None
            df4 = None

            months_processed = months_processed + 1
            print('Successfully scraped !!' )

        except FileNotFoundError:
            print('FileNotFoundError: Payslip doesnot exist')
            continue

        except Exception as e:
            print('Ran into error')
            print(str(e))
            failed_to_parse.append(file_name)
            continue

        except:
            print('Some other error')
            failed_to_parse.append(file_name)
            continue


print('\n---------------------------------------------------------------------------------')
print('\nFailed to parse files : ')
print(failed_to_parse)
print('\nTotal Payslips processed : ' + str(months_processed))

ct = str(datetime.datetime.now())
ct = ct.replace(':','_')
ct = ct.replace('.','_')


all_in_one_excel_path = current_directory + '\payslips_scraped_all_cuts' + '_' + emp_id  + '_' + ct + '.xlsx'

result_raw = pd.concat(pslips)
result_raw['Amount'] = result_raw['Amount'].str.replace(',','')


try:
    result_raw['Amount'] = pd.to_numeric(result_raw['Amount'])
except Exception as e:
    print('\n Failed to convert Amount to numeric :' + str(e))
result = result_raw[~result_raw['Category'].isin(['NET PAY' , 'GROSS EARNING' , 'GROSS DEDUCTIONS'])]


writer = pd.ExcelWriter(all_in_one_excel_path, engine='xlsxwriter')

result.to_excel(writer, sheet_name='Raw Data')
result.groupby([ 'Year_name', 'Month_name', 'Type', 'Category' ])['Amount'].sum().to_excel(writer, sheet_name='Grouped Raw Data')
result.groupby(['Type', 'Category' ])['Amount'].sum().to_excel(writer, sheet_name='Category Aggregated Data')
result.groupby(['Type'])['Amount'].sum().to_excel(writer, sheet_name='Type Aggregated Data')

writer.close()

print('\nAll sheets combined in Excel data stored at : ' + all_in_one_excel_path )

print('\n-----------DONE--------------\n')
