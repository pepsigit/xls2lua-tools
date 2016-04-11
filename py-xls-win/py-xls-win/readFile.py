import os

# Add Analysis File 
from analysis import *
  
def scanfile( _path ):  

    fileList = []  
  
    files = os.listdir( _path )   
 	
	# Scan Files
    for _file in files:  
        if( os.path.isfile( _path + '/' + _file ) ):   
            fileList.append( _file )
   
	# Begin Analysis Xls File
    for _filename in fileList: 
		function_proto( _filename, 'D:/py-xls-win/lua')    
  
if __name__ == '__main__':  
    scanfile( 'D:/py-xls-win/xls')   
