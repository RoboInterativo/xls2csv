#!/usr/bin/python3
from openpyxl import load_workbook

import os

def convert_file(file):
    current_path=( os.path.dirname (os.path.abspath(__file__)) )
    csv_file =file.replace('xlsx','csv')
    f2=open(  os.path.join(current_path, 'out',csv_file,)  ,'w' )
    wb = load_workbook(filename = os.path.join(current_path,'in' ,file ) )
    sheet_ranges = wb[ wb.sheetnames[0]  ]
    values= sheet_ranges.values
    
    for row in  values:
       csv=[]
       for col in row:
          csv.append(col)
       my_string = ','.join(map(str, csv))
       f2.write( my_string+'\n')
    f2.close()
    wb.close()
    os.rename( os.path.join(current_path,'in' ,file ), os.path.join(current_path,'in_arc' ,file )  )
def main():
    files=[]
    current_path=( os.path.dirname (os.path.abspath(__file__)) )
    f2=open(os.path.join(current_path,'in')) ,file.rstrip('.xlsx' )+'.csv','w')
    csv=''
    for f in os.listdir('./in/'):
        
        if f.endswith('.xlsx'):
            files.append(f)
    print  (files)
    print (os.path.join(current_path,'in' ,files[0] ) )
    for file in files:
        convert_file(file)
    #print(sheet_ranges['D18'].value)

if __name__ == "__main__":
    main()
