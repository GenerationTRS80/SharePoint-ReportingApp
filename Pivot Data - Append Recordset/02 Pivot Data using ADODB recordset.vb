Private Function aaCreate_SQL_Join_ADODB_Recordset(xlWrkSht_Button As Excel.Worksheet, _
                            dStartDate As Date, _
                            Optional iNumberWeeks_Ahead As Integer = 5, _
                            Optional sSelect_Report As String = "Pivot", _
                            Optional sTarget_WorksheetName As String = "Pivot_WeeklyHours", _
                            Optional sSet_NameRange_TargetWorksheet As String = "Pivot_Hours", _
                            Optional sPivot_Table As String = "Employee_Opportunity", _
                            Optional sRelated_Table_01 As String = "Employee", _
                            Optional sRelated_Table_02 As String = "Opportunity", _
                            Optional sRelated_Table_03 As String = "Client") As Boolean
                         
 '------------------------------------------------------------------------------------------------------------
 '
 '    1) Automatically write SQL statement to Join Tables. 
 '          The tables that are join are name ranges that contain
 '          data from the downloaded tables data out of the SharePoint PreSales DB
 '    2) Create a ADODB recordset from that SQL statement
 '    3) Write that recordset to a worksheet using ADODB recordset object
 '    4) Take arguments for Worksheet, Date and Integer count
 '

 'Local Variables
  Dim dWeekDay As Date
  Dim sConnString As String
  Dim sSQL As String
  Dim sSQL_PivotTable As String
  Dim sSQL_Related01 As String
  Dim sSQL_Related02 As String
  Dim sSQL_Related03 As String
  
  Dim sAddress As String
  Dim sfilepath As String
  Dim lRowCount As Long
  Dim lRowCount_RecordSet As Long
  Dim lColumnCount As Long
  
 'Copy Recordset to Spreadsheet arguments
  Dim lStartRow As Long
  Dim lStartColumn As Long
  Dim lCursorRow As Long
  Dim lCursorColumn As Long
  Dim bHeaderInclude As Boolean
  Dim bAddFilter As Boolean

  
 'Set workbook objects
  Dim xlWrkBk_Forecast As Excel.Workbook
  Dim rngCell_Table As Range
  Dim rngList_Tables As Range
  Dim xlName As Excel.Name
  
 'ADO Objects
  Dim rsPivot As ADODB.Recordset
  Dim rsDataForecast As ADODB.Recordset
  Dim rsIssueRisk As ADODB.Recordset
  Dim Cmd As ADODB.Command
  Dim Conn As ADODB.Connection
  Dim Rec As ADODB.Record
  
 'ADO Objects NOT instantiated
  Dim rsClone_CopyToSpreadsheet As ADODB.Recordset

  
 'Instantiate objects
  Set rsPivot = New ADODB.Recordset
  Set rsDataForecast = New ADODB.Recordset
  Set rsIssueRisk = New ADODB.Recordset
  Set Cmd = New ADODB.Command
  Set Conn = New ADODB.Connection
  
 'Set Constant
  Const FIELDNAME_PIVOTTABLE_KEY = "Employee Opp Key"
      
  On Error GoTo ProcErr
  
 'Set default value for function
  aaCreate_SQL_Join_ADODB_Recordset = False

  
 'Get workbook name
  Set xlWrkBk_Forecast = xlWrkSht_Button.Parent
  
 'Get list of tables
  Set rngList_Tables = xlWrkBk_Forecast.Names("PDL_ListTables").RefersToRange
  
 'Get file path
  sfilepath = xlWrkBk_Forecast.FullName
  
 'Open Connection
  sConnString = PubFN_Excel_ConnectionString(sfilepath)


 'Get the address from Named Ranges for the listed tables
 '      Pivot Table (employee_opportunity) and related Tables
 '      Related Table 1 (employee)
 '      Related Table 2 (opportunity)
 '      Related Table 3 (client)
 
 
 'Get Pivot Table address from name range
  sAddress = xlWrkBk_Forecast.Names("tbl" & sPivot_Table).RefersToRange.Address(False, False)
  sSQL_PivotTable = "[" & sPivot_Table & "$" & sAddress & "]"
  
 'Get Related Table 01 address from name range
  sAddress = xlWrkBk_Forecast.Names("tbl" & sRelated_Table_01).RefersToRange.Address(False, False)
  sSQL_Related01 = "[" & sRelated_Table_01 & "$" & sAddress & "]"
  
 'Get Related Table 02 address from name range
  sAddress = xlWrkBk_Forecast.Names("tbl" & sRelated_Table_02).RefersToRange.Address(False, False)
  sSQL_Related02 = "[" & sRelated_Table_02 & "$" & sAddress & "]"
  
 'Get Related Table 03 address from name range
  sAddress = xlWrkBk_Forecast.Names("tbl" & sRelated_Table_03).RefersToRange.Address(False, False)
  sSQL_Related03 = "[" & sRelated_Table_03 & "$" & sAddress & "]"
  
 'Open Connection
  Conn.Open sConnString
      
 '>> Instantiate  rsPUBLIC_Pivot <<
  Set rsPUBLIC_Pivot = New ADODB.Recordset
  Set rsDataForecast = New ADODB.Recordset
  Set rsIssueRisk = New ADODB.Recordset

 
 'Disconnect the PUBLIC recordset us cursor location client
  rsPUBLIC_Pivot.CursorLocation = adUseClient
  rsDataForecast.CursorLocation = adUseClient
  rsIssueRisk.CursorLocation = adUseClient
                
       
 ' ************ Join Table Recordset / Pivot True or Do NOT Pivot False  *************

 ' Below create recordset rsPivot from the function FN_Write_SQL Pivot   
  Select Case sSelect_Report
  
    'Report worksheet: Pivot_WeeklyHours
    Case "Pivot"
  
     '*** Pivot recordset ***
     
     'Set recordset to returned value from Command Object
      With rsPivot
          
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            
      End With
      

     'Set Row Count to 0
      lRowCount = 0
       
     '---------------------------------------------------------------------------------------------------------
     '  Create a recordset for each week and append them together into one recordset
     '  NOTE: The recordset append function is used due to the Union SQL does not work for Excel
     
      For i = 0 To iNumberWeeks_Ahead
      
     
        'Set weekday
        dWeekDay = DateAdd("d", 7 * i, dStartDate)
             
            
        '------------------------------------------------------------------------------------------------------
        '
        '     Create a recordset that pivots the weekday from a column to a row  then copy it to a worksheet
        '     NOTE: Use the FN_Write_SQL_Pivot to create the SQL Statement
        '
            
        '                   >>>>>>>>>>> Write SQL statement <<<<<<<<<<<<<<
        sSQL = FN_Write_SQL_Pivot(dWeekDay, WorksheetFunction.WeekNum(dWeekDay), _
                                                   FIELDNAME_PIVOTTABLE_KEY, _
                                                   sSQL_PivotTable, _
                                                   sSQL_Related01, _
                                                   sSQL_Related02, _
                                                   sSQL_Related03)
         

        'Set command object
        With Cmd
            
          .ActiveConnection = Conn
          .CommandText = sSQL
          .CommandType = adCmdText
          .CommandTimeout = 6  'NOTE: if user isn't logged into PreSale DB ie. have Forecast Tool website open. Then app will timeout
            
        End With
                
        
        'Execute command object
        Set rsPivot = Cmd.Execute
            
        'Count the number of rows and columns in the record set
        Do While Not rsPivot.EOF
              
            lRowCount_RecordSet = lRowCount_RecordSet + 1
            rsPivot.MoveNext
        Loop
              
        rsPivot.MoveFirst
               
        ' ***** Append rsPivot to rsPUBLIC_Pivot ******
        If FN_Append_Recordset_rsPUBLIC_Pivot(rsPivot, "RowNumber Pivot WeeklyHours", lRowCount, FIELDNAME_PIVOTTABLE_KEY) = False Then
              
          Debug.Print "**** Error Exited CopyRecordset Sub ******"
          GoTo ProcExit
            
        End If
         
        'Total Rowcount for all recordsets
        lRowCount = lRowCount + lRowCount_RecordSet
               
        'Set lRowCount_RecordSet to 0
        lRowCount_RecordSet = 0
          
      Next i
      
     'Move the first record
      rsPUBLIC_Pivot.MoveFirst
      
     ' *** Clone recordset to rsClone_CopyToSpreadsheet
     'NOTE: The recordset needs to be cloned to pass it to the SUB_PivotSQL_CopyRecordset_to_Spreadsheet
      Set rsClone_CopyToSpreadsheet = rsPUBLIC_Pivot.Clone
      
      
     ' >>>> Set  ARGUMENTS value for copy to spreadsheet <<<
     
     'Include Header True/ Exclude Header False
      bHeaderInclude = True
      
     'Set recordset First Row and the Column to copy recordset
      lStartRow = 3
      lStartColumn = 1
      
     'Set Cursor row and column
      lCursorRow = 4
      lCursorColumn = 2
      
     'Add filter for the Name Range of the copied recordset data in the target worksheet
     'NOTE: if bHeaderInclude = FALSE then a Filter will NOT be added to the name range
      bAddFilter = False
      
      
    'Report worksheet: Data Forecast
    Case "Data Forecast"
  
      '------------------------------------------------------------------------------------------------------
      '
      '     Create a recordset WITHOUT pivoting the data then copy it to a worksheet
      '     NOTE: Use the Function FN_Write_SQL_DataForecast to create the SQL Statement
      '
      
      '                   >>>>>>>>>>> Write SQL statement <<<<<<<<<<<<<<
       sSQL = FN_Write_SQL_DataForecast(dStartDate, WorksheetFunction.WeekNum(dStartDate), _
                                                       6, _
                                                       sSQL_PivotTable, _
                                                       sSQL_Related01, _
                                                       sSQL_Related02, _
                                                       sSQL_Related03)

      'Open Recordset set cursor to client and cursor type to static
       rsDataForecast.Open sSQL, Conn, adOpenStatic, adLockBatchOptimistic
      
       lRowCount_RecordSet = 0
      
      'Count the number of rows and columns in the record set
       Do While Not rsDataForecast.EOF
     
           lRowCount_RecordSet = lRowCount_RecordSet + 1
           rsDataForecast.MoveNext
       Loop
    
       'Move recordset to the first record
       rsDataForecast.MoveFirst
    

      ' *** Clone recordset to rsClone_CopyToSpreadsheet
      'NOTE: The recordset needs to be cloned to pass it to the SUB_PivotSQL_CopyRecordset_to_Spreadsheet
       Set rsClone_CopyToSpreadsheet = rsDataForecast.Clone


      ' >>>> Set ARGUMENTS value for copy to spreadsheet <<<
      'Include Header True/ Exclude Header False
       bHeaderInclude = True
     
      'Set recordset First Row and the Column to copy recordset
       lStartRow = 2
       lStartColumn = 1
      
      'Set Cursor row and column
       lCursorRow = 3
       lCursorColumn = 2
      
      'Add filter for the Name Range of the copied recordset data in the target worksheet
      'NOTE: if bHeaderInclude = FALSE then a Filter will NOT be added to the name range
       bAddFilter = True
        
        
    'Report worksheet: Issue Risk - NextStep
    Case "Issue and Risks"
  
      '------------------------------------------------------------------------------------------------------
      '
      '     Create a recordset WITHOUT pivoting the data then copy it to a worksheet
      '     NOTE: Use the Function FN_Write_SQL_DataForecast to create the SQL Statement
      '
      
      
       sSQL = FN_Write_SQL_IssueRisks(sSQL_PivotTable, _
                                                       sSQL_Related01, _
                                                       sSQL_Related02, _
                                                       sSQL_Related03)


      'Open Recordset set cursor to client and cursor type to static
       rsIssueRisk.Open sSQL, Conn, adOpenStatic, adLockBatchOptimistic
      
       lRowCount_RecordSet = 0
      
      'Count the number of rows and columns in the record set
       Do While Not rsIssueRisk.EOF
     
           lRowCount_RecordSet = lRowCount_RecordSet + 1
           rsIssueRisk.MoveNext
       Loop
    
       rsIssueRisk.MoveFirst
    
      ' *** Clone recordset to rsClone_CopyToSpreadsheet
      'NOTE: The recordset needs to be cloned to pass it to the SUB_PivotSQL_CopyRecordset_to_Spreadsheet
       Set rsClone_CopyToSpreadsheet = rsIssueRisk.Clone


      ' >>>> Set ARGUMENTS value for copy to spreadsheet <<<
      'Include Header True/ Exclude Header False
       bHeaderInclude = True
     
      'Set recordset First Row and the Column to copy recordset
       lStartRow = 2
       lStartColumn = 1
      
      'Set Cursor row and column
       lCursorRow = 3
       lCursorColumn = 2
      
      'Add filter for the Name Range of the copied recordset data in the target worksheet
      'NOTE: if bHeaderInclude = FALSE then a Filter will NOT be added to the name range
       bAddFilter = False
        
  End Select
  

 '    *********** Copy Recordset to Spreadheet *************
  If SUB_PivotSQL_CopyRecordset_to_Spreadsheet(xlWrkBk_Forecast, _
                                                sTarget_WorksheetName, _
                                                rsClone_CopyToSpreadsheet, _
                                                lStartRow, lStartColumn, _
                                                bHeaderInclude, _
                                                6000, 50, _
                                                lCursorRow, lCursorColumn, _
                                                sSet_NameRange_TargetWorksheet, _
                                                bAddFilter) = False Then
 
    Debug.Print "**** Error Exited CopyRecordset Sub ******"
    GoTo ProcExit

  End If

 'Set function to TRUE when all actions are completed
  aaCreate_SQL_Join_ADODB_Recordset = True
  
ProcExit:
  
  'Close Connection object
  Conn.Close
  Set Conn = Nothing
  
  rsPivot.Close
  Set rsPivot = Nothing
  
  rsDataForecast.Close
  Set rsDataForecast = Nothing
  
  rsIssueRisk.Close
  Set rsIssueRisk = Nothing
  
  rsPUBLIC_Pivot.Close
  Set rsPUBLIC_Pivot = Nothing
  
  rsClone_CopyToSpreadsheet.Close
  Set rsClone_CopyToSpreadsheet = Nothing
   
Exit Function


ProcErr:

  Select Case Err.Number
  
    Case 91, 424 'Object not found Note: This occurs on the rsTrackChanges close statement
      'Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
      Resume Next
      
    Case 3021 'No record time has not been entered
      aaCreate_SQL_Join_ADODB_Recordset = False
      MsgBox "Hours have not been entered for the weeks or week you have selected" & vbCrLf & vbCrLf & _
              "You will need to select prior week period", vbOKOnly + vbInformation
      Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
      Resume ProcExit

    Case 3704 'Recordset empty End program to stop more errors
      Resume Next
      
    Case -2147217913, -2147217900, -2147217904 'Error with the Criteria of the expression or SQL statement
      aaCreate_SQL_Join_ADODB_Recordset = False
      Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
      MsgBox "Send email to ITOPursuitsites@atos.net stating there is an error with Forecast Tools Reports" & vbCrLf & " With error # " & Err.Number & vbCrLf & "Send email to ITOPursuitsites@atos.net", vbExclamation + vbOKOnly
      Debug.Print sSQL
      Resume ProcExit
    
    Case Else
      Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
      MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical
      Resume ProcExit
    
  End Select
   
Resume ProcExit

End Function