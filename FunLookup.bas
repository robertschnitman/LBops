'=====================================================================================
' REPO: LBops
' MODULE: FunLookup.bas
' DESCRIPTION: Simplified lookup functions.
'
' LIST OF FUNCTIONS:
' ---CallFuns---
' INDEXC()
' XMATCH()
' VLOOKUPC()
' ---Main---
' FLOOKUP()
' INDEXH()
' INDEXM()/INDEXMATCH()
' SMATCH()
'=====================================================================================

'=====================================================================================
' Function: INDEXC()
' DESCRIPTION: Easier call to INDEX within Developer mode.

Function INDEXC(DataRange as Range, RowNum as Integer, Optional ColNum as Integer)

	fa = createUnoService("com.sun.star.sheet.FunctionAccess")
	
	INDEXC = fa.callFunction("INDEX", Array(DataRange, RowNum, ColNum))

End Function
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

Function VLOOKUPC(IDLookup, DataRange, ColLookup, Optional ApproxMatch as Boolean)

	If IsMissing(ApproxMatch) Then
	
		ApproxMatch = False
		
	End If
		

	fa = createUnoService("com.sun.star.sheet.FunctionAccess")
	
	VLOOKUPC = fa.callFunction("VLOOKUP", Array(IDLookup, DataRange, ColLookup, ApproxMatch))

End Function
'=====================================================================================

'=====================================================================================
' Function: FLOOKUP()
' DESCRIPTION: "Flexible Lookup", a simpler Index-Match formula.
Function FLOOKUP(IDLookup, DataRange as Variant, NamesPattern, NamesRange as Variant)
    
    ColNum = XMATCH(NamesPattern, NamesRange)
    
    FLOOKUP = VLOOKUPC(IDLookup, DataRange, ColNum, False)

End Function
'=====================================================================================

'=====================================================================================
' FUNCTION: INDEXH()
' DESCRIPTION: Get Header value in range.
Function INDEXH(DataRange As Range, DataHeader As Range, StringPattern As String)
	
	' Regular INDEX calculation with XMATCH
	xm = XMATCH(StringPattern, DataHeader)		
	
	' Result
	INDEXH = INDEXC(DataRange, 1, xm)
 
End Function
'=====================================================================================

'=====================================================================================
' Function: INDEXMATCH()/INDEXM()
' Description: Simplification of index-matching [INDEX(..., MATCH(...), MATCH(...))].

Function INDEXMATCH(DataRange, LookupVal1, MatchRange1, LookupVal2, MatchRange2)	
	
	RowNum = XMATCH(LookupVal1, MatchRange1)
	
	ColNum = XMATCH(LookupVal2, MatchRange2)
	
	INDEXMATCH = INDEXC(DataRange, RowNum, ColNum)

End Function

'INDEXMATCH() Synonym
Function INDEXM(DataRange, LookupVal1, MatchRange1, LookupVal2, MatchRange2)

	INDEXM = INDEXMATCH(DataRange, LookupVal1, MatchRange1, LookupVal2, MatchRange2)

End Function
'=====================================================================================

'=====================================================================================
' FUNCTION: SMATCH()
' DESCRIPTION: Simplified XMATCH().
Function SMATCH(StringPattern As String, DataRange As Range)
	
	SMATCH = XMATCH(StringPattern, DataRange, 0)
    
End Function
'=====================================================================================