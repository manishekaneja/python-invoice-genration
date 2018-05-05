import pandas as pd
import numpy as np
import os
import argparse

parser = argparse.ArgumentParser(description='Process some integers.')
parser.add_argument('--p',default='/home/manish/Desktop/Task1/',
                    help='path to files')

args = parser.parse_args()

array_data=[]

path_dict={}
path_dict['invoice_folder']=args.p
path_dict['softData']=args.p+'SoftDataUpload.xlsx'
path_dict['COD']=args.p+'COD.xlsx'
path_dict['PPD']=args.p+'PPD.xlsx'

def getKey(x):
    try:
        softData=pd.read_excel(path_dict[x])
        check = False
        counter = 0
        while (not check) and counter < len(softData):        
            if (softData.loc[counter][1] == "Unused"):
                check=True
                break;
            counter+=1
        strg = softData.loc[counter][0]
        softData.loc[counter][1]="Used"    
        writer = pd.ExcelWriter(path_dict[x], engine='xlsxwriter')
        softData.to_excel(writer, sheet_name=x)
        return strg
    except:
        print('CODE LIST may be fully used')
        exit()

class SoftFormat:
    def __init__(self,mat):
        self.order_number=mat[8,1]
        self.product_mode=mat[5,3].split("-")[0].strip()
        self.consignee=mat[11,2]
        self.awb_number=getKey("COD" if "COD" in self.product_mode else "PPD")
        self.consignee_address=mat[12,2]
        self.destination_city=mat[12,2].split(",")[-1].strip()
        #self.pincode=mat[15,3]
        self.pincode="N.A."
        self.state=mat[14,3]
        self.mobile_number=mat[13,2].split(":")[1].strip()
        self.item_desc=mat[20,1]
        self.pieces=mat[20,6]
        self.declared_value=mat[26,9]
        self.collectable_value=self.declared_value if self.product_mode == "COD" else 0
        self.invoice_number=mat[5,1]
        self.invoice_date=mat[5,8]
        self.seller_gstin=mat[2,1].split("-")[1].strip()
        self.gst_tax_name="HR "+str(mat[25,8]) 
        self.gst_tax_base=mat[20,9]
        self.gst_tax_total=mat[25,9]
    def get_array(self):
        return [self.awb_number,self.order_number,self.product_mode,self.consignee,
                self.consignee_address,self.destination_city,self.pincode,self.state,
                self.mobile_number,self.item_desc,self.pieces,self.collectable_value,
                self.declared_value,self.invoice_number,self.invoice_date,self.seller_gstin,
                self.gst_tax_name,self.gst_tax_base,self.gst_tax_total]

print("READING INVOICE.....")
for filename in os.listdir(path_dict['invoice_folder']):
    if 'invoice' in filename:
        print("--> \t "+filename)
        xl_file = pd.read_excel(path_dict['invoice_folder']+filename)
        x=np.asarray(xl_file)
        array_data.append(SoftFormat(x[:,0:11]))
        array_data.append(SoftFormat(x[:,12:-1]))

softData=pd.read_excel(path_dict['softData'])
print("Extracting Data.....")

for x in array_data:
    softData.loc[len(softData) or 0]=x.get_array()

writer = pd.ExcelWriter(path_dict['softData'], engine='xlsxwriter')
softData.to_excel(writer, sheet_name='Sheet1')
writer.save()
print("Done.......")

