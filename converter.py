#!/usr/bin/env python
#from openpyxl import load_workbook

import os
path_in='in'
path_in_arc='in_arc'
path_out='out'
temp='tmp'


def convert_file(file):
    current_path=( os.path.dirname (os.path.abspath(__file__)) )
    csv_file =file.replace('.xlsx','')
    input_file= os.path.join(current_path,path_in ,file)
    #output_file= os.path.join(current_path,path_out ,csv_file)
    temp_dir=  os.path.join(current_path ,temp)
    
    shell_command_str = 'python xlsx2csv.py  {0} {1} --all'.format(input_file, temp_dir)
    #print shell_command_str
    os.system ( shell_command_str )
    os.remove(   os.path.join(current_path,path_out  ,file.replace('.xlsx','.csv' )   )) 
    for file_name in os.listdir( temp_dir ):
       os.rename(  os.path.join (temp_dir ,file_name ) , os.path.join(current_path, path_out , csv_file+'_'+file_name )  )

    #os.rename( os.path.join(current_path,path_in ,file ), os.path.join(current_path,path_in_arc ,file )  )
     
def main():
    files=[]
    current_path=( os.path.dirname (os.path.abspath(__file__)) )

    for file_name in os.listdir( os.path.join(current_path, path_in )):
        f2=open( os.path.join(current_path, path_out , file_name.rstrip ('.xlsx' )+'.csv'  ), 'w'  )
        csv=''
        if file_name.endswith('.xlsx'):
            convert_file(file_name)


if __name__ == "__main__":
    main()
