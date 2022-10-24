'=====================================================================================
' REPO: LBops
' MODULE: FunLookup.bas
' DESCRIPTION: Simplified lookup functions.
'
' LIST OF FUNCTIONS:
' SMATCH()/MATCHS()
' INDEXH()
' INDEXMATCH()/INDEXM
'=====================================================================================

'=====================================================================================
' Function: SMATCH()/MATCHS()
' Description: Determine which cell in a range matches a given pattern.

Function SMATCH(pattern As String, myrange As Range)

	fa = createUnoService("com.sun.star.sheet.FunctionAccess")

	' MATCH() pattern-matching in LibreOffice uses zero instead of 2 like in Excel.
	SMATCH = fa.callFunction("MATCH", Array(pattern, myrange, 0))
    
End Function

'SMATCH() Synonym
Function MATCHS(pattern As String, myrange As Range)
	
	MATCHS = SMATCH(pattern, myrange)
    
End Function

'=====================================================================================

'=====================================================================================
' FUNCTION: INDEXH()
' DESCRIPTION: Get Header value in range.
Function INDEXH(dataref As range, dataheader As range, pattern As String)

    fa = createUnoService("com.sun.star.sheet.FunctionAccess")
	
	' Regular INDEX calculation with XMATCH
	xm = SMATCH(pattern, dataheader)
		
	output = fa.callFunction("INDEX", Array(dataref, 1, xm))
    
    INDEXH = output    
 
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