'=====================================================================================
' REPO: LBops
' MODULE: FunDates.bas
' DESCRIPTION: Functions for parsing dates.
'
' LIST OF FUNCTIONS:
' WEEKDAYNAME()
' YMD()
' MDY()
' DMY()
'=====================================================================================

'=====================================================================================
' Function: WEEKDAYNAME()
' Description: Outputs the name of the weekday for a given date.

Function WEEKDAYNAME(d As Date)

    wday = Weekday(d, vbSunday) ' 1 = Sunday
    
    Select Case wday
    
        Case 1
        
            result = "Sunday"
            
        Case 2
        
            result = "Monday"
            
        Case 3
            
            result = "Tuesday"
            
        Case 4
        
            result = "Wednesday"
            
        Case 5
        
            result = "Thursday"
            
        Case 6
        
            result = "Friday"
            
        Case 7
        
            result = "Saturday"
            
    End Select
    
    WEEKDAYNAME = result

End Function

'=====================================================================================
' Function: YMD()
' Description: Formats a date value into the ISO standard format ("yyyy-mm-dd").

Function YMD(d As Date)

    YMD = Format(d, "yyyy-mm-dd")

End Function

'=====================================================================================
' Function: MDY()
' Description: Formats a date value into the month-day-year order ("mm/dd/yyyy").

Function MDY(d As Date)

    MDY = Format(d, "mm/dd/yyyy")

End Function

'=====================================================================================
' Function: DMY()
' Description: Formats a date value into the day-month-year order ("dd/mm/yyyy").

Function DMY(d As Date)

    DMY = Format(d, "dd/mm/yyyy")

End Function