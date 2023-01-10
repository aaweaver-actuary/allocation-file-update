import pandas as pd
import numpy as np

#Change these to ask for input later:
cy = 2022 #Change to ask for input
cm = 2 #Change to ask for input
#Ask for Quarter, too

#Empty output table to concat in loop:
output = pd.DataFrame(columns=['ay','lob','cum_rpt_loss','cum_paid_dcce','ep','sel_loss_ratio','sel_dcce_ratio'])

#Location of all link ratio analyses for 4Q2021:
foldername = r'O:\STAFFHQ\SYMDATA\Actuarial\Reserving Applications\IBNR Allocation\4Q2021 Analysis\CSU'

#Loop for each line:
#Change to ask for input later
filelist = [r'\CSU Other Liability - Occurrence 4Q2021.xlsm',r'\CSU Allied 4Q2021.xlsm',r'\CSU Boiler & Machinery 4Q2021.xlsm']
lob_list = ['OL-Occ','Allied','B&M']
for file in filelist:
    
    #Actual filepath:
    filename = foldername+file

    #Read in the DashUpload sheet:
    df = pd.read_excel(filename, sheet_name = 'DashUpload')

    #Cumulative reported loss with AY and Value columns:
    cum_rpt_loss = df[(df['item_type']=='reported_loss') & (df['item_sub_type']=='cumulative')][['item_row_lookup','Value']].reset_index(drop=True)

    #Cumulative paid dcce:
    cum_paid_dcce = df[(df['item_type']=='paid_dcce') & (df['item_sub_type']=='cumulative')][['Value']].reset_index(drop=True)

    #Earned premium:
    earn_prem = df[(df['item_type']=='premium') & (df['item_sub_type']=='earned')][['Value']].reset_index(drop=True)

    #Selected ultimate reported loss ratio:
    loss_ratio = df[(df['item_type']=='reported_loss') & (df['item_sub_type']=='selected_ult_loss_ratio')][['Value']].reset_index(drop=True)

    #Selected ultimate DCCE ratio:
    dcce_ratio = df[(df['item_type']=='paid_dcce') & (df['item_sub_type']=='selected_ult_loss_ratio')][['Value']].reset_index(drop=True)

    #Output for this lob:
    d = {'ay':cum_rpt_loss['item_row_lookup'],'lob':[lob_list[filelist.index(file)]]*len(cum_rpt_loss.index),'cum_rpt_loss':cum_rpt_loss['Value']}
    lob_output = pd.DataFrame(d)
    lob_output['cum_paid_dcce'] = cum_paid_dcce
    lob_output['ep'] = earn_prem
    lob_output['sel_loss_ratio'] = loss_ratio
    lob_output['sel_dcce_ratio'] = dcce_ratio
                      
    #Accumulating Output Table:
    output = pd.concat([output,lob_output]) #Update this to include all items
    
#Read out to check:
output.to_excel('Output_Check_4.xlsx',sheet_name='Table')
