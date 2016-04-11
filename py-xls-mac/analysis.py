import xdrlib
import xlrd
import os

t_empty 	= 0
t_string 	= 1
t_number 	= 2
t_data 		= 3
t_boolean 	= 4
t_error 	= 5
 

def function_proto( _file, path ):
	
	filename = _file[ 0 : len(_file) - 4 ]
	
	if( -1 == cmp( os.path.splitext(_file)[1], '.xls') ):
		return
		
	data = xlrd.open_workbook( 'xls/' + _file )
	
	#print( 'FileName = ', filename )

	#################

	tables = data.sheet_names()
	tablesnum = len(tables)
	#print( 'tablesnum = ', tablesnum, tables )

	#################
	
	output = open( '/' + path + '/' + filename + '.lua', 'wb+' )

	output.write( filename + ' = \n' )
	output.write( '{ \n' )
	
	#################

	for _table in range( tablesnum ):
	
		# Get Table By Index
		table = data.sheet_by_index( _table )
		#print( 'table - sheetname = ', table.name,table.cell_value( 0, 0 ) )
		
		if( -1 == cmp(table.cell_value( 0, 0 ), 'id') ):
			output.seek(0)
			output.truncate()
			output.write('The First Col Of The Sheet Named [ \"' + table.name + '\" ] Must Be \'id\'.')
			return
		
		output.write( '\t' + table.name + ' =\n' )
		output.write( '\t{\n' )
		
		# Get Current Table Row
		nrows = table.nrows
		#print( 'nrows = ', nrows )
		
		# Get Current Table Col
		ncols = table.ncols
		#print( 'ncols = ', ncols )

			
		#################
		
		for _row in range( nrows - 1 ):
			
			# First Col Must Be 'id'
			output.write('\t\t[')
			output.write( str( int( table.cell_value( _row + 1, 0 ) ) ) )
			output.write('] =\n')
			output.write('\t\t{\n')
							
			for _col in range( ncols - 1 ):
				
				# Some Kinds Of Data Type.
				#print( table.cell( _row + 1, _col + 1 ).ctype )
				if( t_number == table.cell( _row + 1, _col + 1 ).ctype ):
					# Num Type.
					#print( table.cell( _row + 1, _col + 1 ).ctype )
					output.write('\t\t\t')
					output.write( str( table.cell_value( 0, _col + 1 ) ) )
					output.write(' = ')
					if( 0 == table.cell_value( _row + 1, _col + 1 ) % 1 ):
						output.write( str( int( table.cell_value( _row + 1, _col + 1 ) ) ) )
					else:
						output.write( str( table.cell_value( _row + 1, _col + 1 ) ) )
					output.write(',\n')
					
				if( t_string == table.cell( _row + 1, _col + 1 ).ctype ):
					
					# Nest Table.
					
					symbol_end_count = table.cell_value( _row + 1, _col + 1 ).count( '#' )
					if( 0 != symbol_end_count ):
						
						_len = len( table.cell_value( _row + 1, _col + 1 ) )
						if( '#' != table.cell_value( _row + 1, _col + 1 )[ _len - 1 ] ):
							output.seek(0)
							output.truncate()
							output.write('The [' + ( '%s' % (_row + 2) ) + ', ' + ( '%s' % (_col + 2) ) + '] Of The Sheet Named [ \"' + table.name + '\" ] , The Last Symbol Must Be \'#\'.')
							return
						else:
							output.write('\t\t\t')
							output.write( str( table.cell_value( 0, _col + 1 ) ) )
							output.write(' = \n')
							output.write('\t\t\t{\n')
							_err = nest_table( table.cell_value( _row + 1, _col + 1 ), symbol_end_count, output )
							output.write('\t\t\t},\n')
					else:	
					# Normal String Type.
						#print( table.cell( _row + 1, _col + 1 ).ctype )
						output.write('\t\t\t')
						output.write( str( table.cell_value( 0, _col + 1 ) ) )
						output.write(' = \"')
						output.write( str( table.cell_value( _row + 1, _col + 1 ).encode('utf-8') ) )
						output.write('\",\n')

			output.write('\t\t},\n')
		output.write( '\t},\n ' )
	output.write( '}' )
	output.close()
	
	
def nest_table( str, s_e_count, output ):
	
	s_e_pos = 0
	for i in range( s_e_count ):
		output.write('\t\t\t\t[')
		output.write( '%s' % (i+1) )
		output.write('] =\n')
		output.write('\t\t\t\t{\n')
		_pos = str.find( '#', s_e_pos, len( str ) - 1 )
		_substr = str[ s_e_pos : _pos ]
		#print( '_substr = ', _substr )
		_sub_div_count = _substr.count( '@' )
		
		# Before '#' Not Have '@'
		if( 0 == _sub_div_count ):
			output.write('\t\t\t\t\t[')
			output.write( '%s' % (1) )
			output.write('] = ')
			output.write( _substr[ 0 : len( _substr ) ] )
			output.write(',\n')
		else:
			s_d_pos = 0
			for j in range( _sub_div_count ):
				__pos = _substr.find( '@', s_d_pos, len( _substr ) - 1 )
				__value = _substr[ s_d_pos : __pos ]
				output.write('\t\t\t\t\t[')
				output.write( '%s' % (j+1) )
				output.write('] = ')
				output.write( __value )
				output.write(',\n')
				s_d_pos = __pos + 1
			output.write('\t\t\t\t\t[')
			output.write( '%s' % (_sub_div_count + 1) )
			output.write('] = ')
			output.write( _substr[ s_d_pos : len( _substr ) ] )
			output.write(',\n')	
			
		output.write('\t\t\t\t},\n')
		
		s_e_pos = _pos + 1
