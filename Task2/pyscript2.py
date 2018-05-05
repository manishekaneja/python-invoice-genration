import pandas as pd
import numpy as np
import os
import argparse

parser = argparse.ArgumentParser(description='Process some integers.')
parser.add_argument('--p',default='/home/manish/Desktop/Task2/',
                    help='path to files')

args = parser.parse_args()
paths={}
paths['SoftFile']=args.p+'orders.xlsx'
paths['s']=args.p+'invoice.xlsx'
df=pd.read_excel(paths['SoftFile'])
data={}
for name in df['Name']:
    data[name]={}
    for head in df:
        if 'Lineitem' in head:
            data[name]['line']=[]
        else:
            data[name][head]={}
for row in range(len(df)):
    nomore=True
    for x in df:
        if  (not df.isnull().loc[row][x]):
            if 'Lineitem' in x:
                if nomore:
                    nomore=False
                    data[df.loc[row][0]]['line'].append([df.loc[row]['Lineitem name'],df.loc[row]['Lineitem price'],df.loc[row]['Lineitem quantity']])
            else:
                data[df.loc[row][0]][x]=df.loc[row][x]
data['PCHTST01']
in_sample=pd.read_excel(paths['s'])
x=np.asarray(in_sample)
final_bill=[]
x=x[:,0:12]
import math as n
for index in range(4):
        x[20+index,1]=n.nan
        x[20+index,0]=n.nan
        x[20+index,6]=n.nan
        x[20+index,5]=n.nan
        x[20+index,7]=n.nan
        x[20+index,8]=n.nan
        x[20+index,9]=n.nan
def cret(y,x):
    x[2,6]="Email-"+y['Email']
    x[5,8]=str(y['Created at'].date())
    x[5,1]=y['Name']   
    x[8,1]=y['Name']
    x[5,3]=y['Payment Method']
    x[11,2]=y['Billing Name']
    x[12,2]=y['Shipping Address1']
    x[13,2]="Mobile Number : "+y['Shipping Phone']
    x[14,3]=y['Shipping City']
    for index,ele in enumerate(y['line']):
        x[20+index,1]=ele[0]
        x[20+index,0]=index
        x[20+index,6]=ele[2]
        x[20+index,5]=6101
        x[20+index,7]=ele[1]
        x[20+index,8]="Nos"
        x[20+index,9]=ele[2]*ele[1]
    x[25,9]=y['Taxes']
    x[26,9]=y['Total']
    x[28,2]=y['Total']
    return x
for bill in data:
    final_bill.append(pd.DataFrame(cret(data[bill],x.copy())))
for x in range(int(len(final_bill)/2)):
    writer = pd.ExcelWriter(args.p+'invoice_'+str(x)+'.xlsx', engine='xlsxwriter')
    pd.DataFrame(np.concatenate((final_bill[x*2],final_bill[((x*2)+1)]),axis=1)).to_excel(writer, sheet_name='Sheet1')
    writer.save()
