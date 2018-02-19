Public Sub SUB_MainSheet_SetListBoxObject_toMonday()
    
    Dim dDate As Date
    Dim dMondayDate As Date
    
   'Get todays date
    dDate = Date

    If WeekDay(dDate, vbSunday) = 2 Then
    
         dMondayDate = Date
        
    Else
         'Find the Monday date for today date
         dMondayDate = DateAdd("d", WeekDay(dDate, vbTuesday) * -1, dDate)
    
    End If

    With Sheet1.CboListDates

        .Text = dMondayDate
        '.DropDown

     End With


End Sub
