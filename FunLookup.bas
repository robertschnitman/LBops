'=====================================================================================
' REPO: LBops
' MODULE: FunLookup.bas
' DESCRIPTION: Simplified lookup functions.
'
' LIST OF FUNCTIONS:
' XMATCH()
' VLOOKUPC()
' SMATCH()/MATCHS()
' INDEXH()
' INDEXMATCH()/INDEXM
'=====================================================================================

'=====================================================================================
' Function: XMATCH()
' DESCRIPTION: Replication of Excel's XMATCH.
Function XMATCH(StringPattern As String, DataRange As Range, Optional MatchType)

	' Handling defaults
	If IsMissing(MatchType) Then
	
		MatchType = 0
		
	End If

	' To call worksheet functions
	fa = createUnoService("com.sun.star.sheet.FunctionAccess")

	' MATCH() StringPattern-matching in LibreOffice uses zero instead of 2 like in Excel.
	XMATCH = fa.callFunction("MATCH", Array(StringPattern, DataRange, MatchType))

End Function
'=====================================================================================

'=====================================================================================
' Function: VLOOKUPC()
' DESCRIPTION: Easier call to VLOOKUP within Developer mode.

Function VLOOKUPC(IDLookup, DataRange, ColumnLookup, Optional ApproxMatch as Boolean)

	If IsMissing(ApproxMatch) Then
	
		ApproxMatch = False
		
	End If
		

	fa = createUnoService("com.sun.star.sheet.FunctionAccess")
	
	VLOOKUPC = fa.callFunction("VLOOKUP", Array(IDLookup, DataRange, ColumnLookup, ApproxMatch))

End Function
'=====================================================================================

'=====================================================================================
' Function: SMATCH()/MATCHS()
' Description: Determine which cell in a range matches a given StringPattern.

Function SMATCH(StringPattern As String, DataRange As Range, Optional MatchType)

	If IsMissing(MatchType) Then
	
		MatchType = 0
		
	End If

	fa = createUnoService("com.sun.star.sheet.FunctionAccess")

	' MATCH() StringPattern-matching in LibreOffice uses zero instead of 2 like in Excel.
	SMATCH = fa.callFunction("MATCH", Array(StringPattern, DataRange, MatchType))
    
End Function

'SMATCH() Synonym
Function MATCHS(StringPattern As String, DataRange As Range)
	
	MATCHS = SMATCH(StringPattern, DataRange)
    
End Function

'=====================================================================================

'=====================================================================================
' FUNCTION: INDEXH()
' DESCRIPTION: Get Header value in range.
Function INDEXH(dataref As range, dataheader As range, StringPattern As String)

    fa = createUnoService("com.sun.star.sheet.FunctionAccess")
	
	' Regular INDEX calculation with XMATCH
	xm = SMATCH(StringPattern, dataheader)
		
	INDEXH = fa.callFunction("INDEX", Array(dataref, 1, xm))        
 
End Function
'=====================================================================================

'=====================================================================================
' Function: INDEXMATCH()/INDEXM()
' Description: Simplification of index-matching [INDEX(..., MATCH(...), MATCH(...))].

Function INDEXMATCH(datarange, lookupval1, range1, lookupval2, range2)

	fa = createUnoService("com.sun.star.sheet.FunctionAccess")
	
	r = SMATCH(lookupval1, range1)
	
	c = SMATCH(lookupval2, range2)
	
	INDEXMATCH = fa.callFunction("INDEX", Array(datarange, r, c))

End Function

'INDEXMATCH() Synonym
Function INDEXM(datarange, lookupval1, range1, lookupval2, range2)

	INDEXM = INDEXMATCH(datarange, lookupval1, range1, lookupval2, range2)

End Function
'=====================================================================================