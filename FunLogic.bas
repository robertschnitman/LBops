'=====================================================================================
' MODULE: FunLogic.bas
' DESCRIPTION: Boolean functions for common scenarios.
'
' LIST OF FUNCTIONS:
' ISLEN0()
' ISERROR()
' IFBLANK()
' SKIPBLANK()
' DOIF()
'=====================================================================================

'=====================================================================================
' Function: ISLEN0()
' Description: Test whether a cell has no characters. Similar to ISBLANK().

Function ISLEN0(cell As String)
    
    ISLEN0 = (Len(cell) = 0)

End Function
'=====================================================================================

'=====================================================================================
' Function: ISERROR()
' Description: Synonym for ISERR for Excel VBA compatibility.

Function ISERROR(cell)

	fa = createUnoService("com.sun.star.sheet.FunctionAccess")
    
    ISERROR = fa.callFunction("ISERR", Array(cell))

End Function
'=====================================================================================

'=====================================================================================
' Function: IFBLANK()
' Description: Similar to IF(), but performs an action depending on whether a cell is blank or not.

Function IFBLANK(cell As String, ValTrue, ValElse)

    If ISLEN0(cell) = True Then
    
        result = ValTrue
        
    Else
    
        result = ValFalse
        
    End If
    
    IFBLANK = result

End Function

'=====================================================================================
' Function: SKIPBLANK()
' Description: Perform an action if a cell is non-blank; otherwise, result blank.

Function SKIPBLANK(cell As String, ValElse)

    If ISLEN0(cell) = True Then
    
        result = ""
        
    Else
    
        result = VALnonblank
        
    End If
    
    SKIPBLANK = result

End Function
'=====================================================================================

'=====================================================================================
' Function: DOIF()
' Description: Perform an action only if a condition is met; otherwise, result blank.

Function DOIF(condition As Boolean, ValTrue)

    If condition = True Then
    
        result = ValTrue
        
    Else
    
        result = ""
        
    End If
    
    DOIF = result

End Function
'=====================================================================================