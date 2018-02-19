Public Sub Button_RunPivot(xlWrkSht_Main As Excel.Worksheet)

  'Excel Objects
   Dim xlWrkBk_Forecast As Excel.Workbook
   Dim xlWrkSht As Excel.Worksheet
   Dim xlCellSelect As Excel.Range
   Dim rngShowReport As Excel.Range

  'Local Variables
   Dim sWorksheet_Name As String
   Dim dStartDate As Date
   Dim iWeeks_Ahead As Integer
   Dim sCboWeeksAhead_ReturnValue As String
   Dim sShowReport As String
   
  On Error GoTo ProcErr

  'Turn Worksheet Protection OFF
   FN_Public_UnProtect_Workbook
   
  'Turn Screen Updating OFF
   With Application
 
    .ScreenUpdating = False
 
   End With
   
  'Set Workbook object from Worksheet object
   Set xlWrkBk_Forecast = xlWrkSht_Main.Parent
   
  'Select worksheet after RunReport
   xlWrkSht_Main.Calculate
   Set rngShowReport = xlWrkBk_Forecast.Names("ShowReport").RefersToRange
   
  'Get the Worksheet name from the ShowReport combobox from the Main page
   sShowReport = rngShowReport.Value

  'Get Worksheet name
   sWorksheet_Name = xlWrkSht_Main.Name
  
  
  '     *** Get Parameters  ***
  'Set Value for the Start Date and the number of weeks ahead to report
  
  ' dStartDate = Sheet1.CboListDates.Value
   dStartDate = xlWrkBk_Forecast.Names("vbParam_Select_StartDate").RefersToRange.Value
   
  'sCboWeeksAhead_ReturnValue = Sheet1.cboWeeksAhead.Value
   sCboWeeksAhead_ReturnValue = xlWrkBk_Forecast.Names("vbParam_Weeks_Ahead").RefersToRange.Value
   
  'Calculate the number of weeks ahead base on the Text String pulled from the name range vbParam_Weeks_Ahead"
   If sCboWeeksAhead_ReturnValue = "Current Week" Then
        
        iWeeks_Ahead = 0
   
   Else
   
        iWeeks_Ahead = CInt(Left(sCboWeeksAhead_ReturnValue, 1))
        
   End If
    
  '>>>>>>>>>>>>>>>>>> RUN Pivot on tblEmployee_Opportunity <<<<<<<<<<<<<<<<<<<
   If aaCreate_SQL_Join_ADODB_Recordset(xlWrkSht_Main, _
                                        dStartDate, _
                                        iWeeks_Ahead, _
                                        "Pivot", _
                                        "Pivot_WeeklyHours", _
                                        "Pivot_Hours") = False Then

        
       'If there is an error with the import then goto procexit
        GoTo ProcExit
        
   End If

  '------------------------- RUN Data Forecast Report -------------------------------
   If aaCreate_SQL_Join_ADODB_Recordset(xlWrkSht_Main, _
                                        dStartDate, _
                                        iWeeks_Ahead, _
                                        "Data Forecast", _
                                        "Data Forecast", _
                                        "Data_Forecast") = False Then
        
       'If there is an error with the import then goto procexit
        GoTo ProcExit

   End If


 '---------------------------------------------------------------------
 '      Select worksheet after report runs
 '      NOTE:   See if the worksheet David L Report can be found/listed.
 '              If it is found then select that worksheet
  For Each xlWrkSht In xlWrkBk_Forecast.Worksheets

     'If the worksheet name exist in the workbook then set the cursor to cell
      If xlWrkSht.Name = sShowReport Then
      
           'Activate worksheet
            xlWrkSht.Activate
             
           'Select cell in worksheet to set cursor
            Set xlCellSelect = xlWrkSht.Cells(1, 5)
            xlCellSelect.Select
         
            Exit For
    
      End If

  Next


ProcExit:

 'Turn Worksheet Protection ON
  FN_Public_Protect_Workbook
  
  
 'Turn Screen Updating ON
  With Application
 
    .ScreenUpdating = True
    
  End With
     
     
 'Refresh all pivot tables
  xlWrkBk_Forecast.RefreshAll


Exit Sub


ProcErr:

  Select Case Err.Number
  
    Case 91, 424 'Object not found Note: This occurs on the rsTrackChanges close statement
      'Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
      Resume Next
    
    Case 94 'Parameter not found
      MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbExclamation & vbCrLf & vbCrLf
      Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
      Stop
      Resume Next
      'Resume ProcExit
      
    Case 3704 'Recordset empty End program to stop more errors
      Resume Next

    Case Else
      Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
      MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical
      Stop
      Resume ProcExit
    
  End Select

Resume ProcExit

End Sub