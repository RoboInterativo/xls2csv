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

    for file_name in os.listdir('./in/'):
        f2=open( os.path.join(current_path,'out'  , file_name.rstrip ('.xlsx' )+'.csv'  ), 'w'  )
        csv=''
        if file_name.endswith('.xlsx'):
            convert_file(file_name)


if __name__ == "__main__":
    main()
