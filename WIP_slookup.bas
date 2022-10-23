' === UNDER CONSTRUCTION - DO NOT USE=== '

' Function: SLOOKUP()
' Description: Lookup a value based on a row value and column name.

Function SLOOKUP(id_lookup As String, column_lookup As String, data_range As Range, Optional column_match_type As Integer)
Attribute SLOOKUP.VB_Description = "Lookup a value by row value and column name."
Attribute SLOOKUP.VB_ProcData.VB_Invoke_Func = " \n20"

	col_begin = LBound(data_range, 2)
	col_count = UBound(data_range, 2)
	
	row_begin = LBound(data_range, 1)
	row_count = UBound(data_range, 1)
	
	' Find column position of column_lookup.
	For i = col_begin to col_count
	
		header = data_range(row_begin)(i)
		
		if column_lookup = header then
		
			c = header
			
			Exit For
			
		End If
		
	Next i
	
	c = SMATCH(column_lookup, headers)
	
	' Find row position of id_lookup
	For i = row_begin to row_count
	
		For j = col_begin to col_count
		
			out = data_range(i)(j)
			
			if id_lookup = out Then
			
				r = i
				
				Exit For
				
			End If
			
		Next j
		
		if id_lookup = out Then
		
			Exit For
			
		End If
		
	Next i	
	
	
	' Pass parameters to INDEX() to get the cross-section.
	result = fa.callFunction("INDEX", Array(data_range, r, c))
	

	SLOOKUP = result
 
    

End Function

