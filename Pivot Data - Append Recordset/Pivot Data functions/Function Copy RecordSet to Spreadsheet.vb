Function SUB_PivotSQL_CopyRecordset_to_Spreadsheet(xlWrkBk_TargetWorkbook As Excel.Workbook, _
                                                sTarget_WorksheetName As String, _
                                                rsSpeadsheet As ADODB.Recordset, _
                                                lStartRow As Long, _
                                                lStartColumn As Long, _
                                                Optional bAddHeader As Boolean = False, _
                                                Optional lClearData_Rows As Long = 1, _
                                                Optional lClearData_Columns As Long = 1, _
                                                Optional lSelectRow As Long = 1, _
                                                Optional lSelectColumn As Long = 1, _
                                                Optional sSet_NameRange_TargetWorksheet As String = "", _
                                                Optional bFilterData As Boolean = False) As Boolean


  '-----------------------------------------------------------------------------------------------------------
  '
  '
  ' CopyRecordset to Spreadsheet function
  '                 ADODB recordset (contains data to be written to worksheet via copyrecordset method)
  '                 Excel worksheet object
  '                    (worksheet object receives the recordset from recordset copy method)
  '                 String (name of worksheet that received the recordset)
  '
  ' Arguments Passed
  '
  ' 1) Write recordset data to worksheet
  '              Arguments:
  '                   a) Set workbook: Excel workbook object xlWrkBk_TargetWorkbook
  '                   b) Set worksheet name: String variable sTarget_WorksheetName
  '                   c) Take recordset: rsSpeadsheet
  '
  ' 2) Select first cell to Copyto recordset data
  '    NOTE: The cell is the firts column and row of recordset data
  '              Arguments:
  '               Selection of cell to copy recordset from the 2 long variables (lStartRow and lStartColumn)
  '
  '
  ' Optional arguments
  '
  ' 1) Create Header (default false)
  '         Write header data if argument bAddHeard is TRUE or don't write header FALSE. FALSE is default value
  '         Arguments: boolean (Write a header from the copied recordset TRUE=write header FALSE= do not write)
  '
  ' 2) Clear data and formatting in worksheet
  '        Set number of rows and columns to be cleared
  '                Select end selection cell from the numeric arguments passed to (lClearData_Rows and lClearData_Columns
  '                      Long (2 arguments to select cells in worksheet to clear data and formatting in copied to worksheet)
  '
  ' 3) Select Cursor cells for the worksheet where the recordset is copied too
  '        Arguments: lSelectRow and lSelectColumn to select cell for worksheet cursor
  '                   Variables 2 long (2 arguments to select cells in worksheet to clear data and formatting in copied to worksheet)
  '
  ' 4) Update Name Range's address for the number of rows and columns for the given recordset returned
  '        Arguments: sSet_NameRange_TargetWorksheet
  '
  ' 5) Set Filter to Header TRUE= Add Filter to header, FALSE (default value) = do not add Filter
  '        Arguments: boolean bFilterData
 
 
 
'************************************************************************************************
'*
'* NOTE: DO NOT USE MS EXCEL's  "Selection" for a substitute for a range of cells
'*          Excel will explicitly instantiate the "Selection" if you do use it
'*            You can't close the instantiate selection
'*


'------- Local Variables ----------
 Dim lNumberFields As Long
 Dim lRowCount_RecordSet As Long
 Dim lColumnCount As Long
 Dim iCol As Integer
 Dim lHeaderRows As Long
 
'------- Excel Objects -------
 Dim appXL As Excel.Application
 Dim xlWrkSht_TARGET As Excel.Worksheet
 Dim xlName As Excel.Name

'---------Range Objects--------
 Dim objCell_1 As Excel.Range
 Dim objCell_2 As Excel.Range
 Dim objRange As Excel.Range
 
 
'Set CopyRecordset to Spreadsheet to TRUE
 SUB_PivotSQL_CopyRecordset_to_Spreadsheet = True
 
'---------Constants------------
 Const ADDITIONAL_FIELD_INCREMENT = 2 'This is required to get he last field in a recordset. Recordset field starts with ZEREO
 Const HEADER_ROWS_ADDED = 1
 
 On Error GoTo ProcErr

'Set target worksheet object
 Set xlWrkSht_TARGET = xlWrkBk_TargetWorkbook.Worksheets(sTarget_WorksheetName)
 
 
'Check for Filter
 
'*** CLEAR ALL THE DATA (Clear Contents) IN THE WORKSHEET ***
'Check that row count and column count are not 1
 If lClearData_Rows <> 1 And lClearData_Columns <> 1 Then
 
    Set objCell_1 = xlWrkSht_TARGET.Cells(lStartRow, lStartColumn)
    Set objCell_2 = xlWrkSht_TARGET.Cells(lStartRow + lClearData_Rows, lStartColumn + lClearData_Columns)
 
    Set objRange = xlWrkSht_TARGET.Range(objCell_1, objCell_2)
    
    objRange.ClearContents
    
 End If
      
'*** Check to see if headers are to be added: AddHeader True/False
 If bAddHeader Then
 
        '*** Copy Field Header into the spreadsheet using recordset fieldnames***
        'NOTE: you need to add 1 to get all fields. Since recordset field start with 0
          
        'Count number of fields
         lNumberFields = rsSpeadsheet.Fields.Count ' + ADDITIONAL_FIELD_INCREMENT

          
        'Copy field names to the first row of the worksheet
         For iCol = lStartColumn To lNumberFields + lStartColumn - 1
         
            xlWrkSht_TARGET.Cells(lStartRow, iCol).Value = rsSpeadsheet.Fields(iCol - lStartColumn).Name
                
         Next

         lHeaderRows = HEADER_ROWS_ADDED
         
     Else
       
       'Zero rows are added when FALSE
        lHeaderRows = 0

 End If


'Check for set Filter in target worksheet, if there is a Filter applied then clear that filtering
 If xlWrkSht_TARGET.FilterMode = True Then

      'Remove Filter
        xlWrkSht_TARGET.ShowAllData

 End If


'--------------------------------------------------------------------------------------------
'            >>>>>>>>>    Copy Recordset to spreadsheet    <<<<<<<<
'
 rsSpeadsheet.MoveFirst
 xlWrkSht_TARGET.Cells(lStartRow + lHeaderRows, lStartColumn).CopyFromRecordset rsSpeadsheet


'  >>>>>   Set Name Range for recordset returned    <<<<<<<
 If Len(sSet_NameRange_TargetWorksheet) <> 0 Then
 
    'Move the first record
     rsSpeadsheet.MoveFirst
     
    'Count the number of rows and columns in the record set
     Do While Not rsSpeadsheet.EOF
       
         lRowCount_RecordSet = lRowCount_RecordSet + 1
         rsSpeadsheet.MoveNext
         
     Loop
     
    '*** Name Range Pivot Hours address
     Set xlName = xlWrkBk_TargetWorkbook.Names(sSet_NameRange_TargetWorksheet)
     
    'Set Name Range address
     xlName.RefersTo = xlName.RefersToRange.Resize(lRowCount_RecordSet + 1, lNumberFields)
      
     
    'Add Filter to given range
     If bFilterData = True Then
     
        xlName.RefersToRange.AutoFilter
 
     End If
      

  End If

 
'Recalculate worksheet
 xlWrkSht_TARGET.Calculate

 
'--------------------------------------------------------------------------------------------
'   NOTE:   You need to check to see if the worksheet is visisible. If NOT visible then
'           formatting and setting of the cursor will *Fail* and cause an error
'

'Format worksheet and set cell for cursor selection
 If xlWrkSht_TARGET.Visible = xlSheetVisible Then
  
  
    'Format sheet
     xlWrkSht_TARGET.Activate
     xlWrkSht_TARGET.Cells.Select
     
      With Selection
          .HorizontalAlignment = xlGeneral
          .VerticalAlignment = xlBottom
          .WrapText = False
          .Orientation = 0
          .AddIndent = True
          .IndentLevel = 0
          .ShrinkToFit = False
          .ReadingOrder = xlContext
          .MergeCells = False
          
      End With

    '*** Set Select cell to set cursor ***
     If lSelectRow <> 1 And lSelectColumn <> 1 Then
    
        Set objCell_1 = xlWrkSht_TARGET.Cells(lSelectRow, lSelectColumn)
        objCell_1.Select
           
     Else
       
       'DO NOT Select the worksheet if set to defualt values
        'Set objCell_1 = xlWrkSht_TARGET.Cells(1, 1)
        'objCell_1.Select
       
     End If


 End If
 

ProcExit:

'Close recordset
 rsSpeadsheet.Close
 Set rsSpeadsheet = Nothing
 
Exit Function


ProcErr:

  Select Case Err.Number
  
    Case 5 'Recordset error
    SUB_PivotSQL_CopyRecordset_to_Spreadsheet = False
    MsgBox " The error # is " & Err.Number & vbCrLf & "Error with CopyRecordset Subroutine " & vbCrLf & "Send email to ITOCostModels with the error description", vbCritical + vbOKOnly
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume ProcExit
    
    Case 9 'Worksheet not found
    SUB_PivotSQL_CopyRecordset_to_Spreadsheet = False
    MsgBox " The error # is " & Err.Number & vbCrLf & "Error with CopyRecordset Subroutine " & vbCrLf & "Send email to ITOCostModels with the error description", vbCritical + vbOKOnly
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume ProcExit
      
    Case 91, 424 'Hourglass Comand
      'Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
      Resume Next
      
    Case 1004
    '   CopyRecordset_to_Spreadsheet = False
    '   MsgBox " Too Many instances of Excel open. Close one or more instances", vbInformation + vbOKOnly
    '   Resume ProcExit
      Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
      Resume Next
      
    Case 3704 'Recordset is already closed
      Resume Next

    Case 3265 'Description of Item Can NOT be found in the recordset
    'If error then set SUB_PivotSQL_CopyRecordset_to_Spreadsheet = False
      SUB_PivotSQL_CopyRecordset_to_Spreadsheet = False
      MsgBox " The error # is " & Err.Number & vbCrLf & "Error with CopyRecordset Subroutine ", vbCritical + vbOKOnly
      Resume ProcExit

    Case -2147467259 'Steam Object can't be read because it is empty
    'If error then set SUB_PivotSQL_CopyRecordset_to_Spreadsheet = False
      SUB_PivotSQL_CopyRecordset_to_Spreadsheet = False
      MsgBox " The error # is " & Err.Number & vbCrLf & "Error with CopyRecordset Subroutine ", vbCritical + vbOKOnly
      Resume ProcExit
      
    Case Else
      SUB_PivotSQL_CopyRecordset_to_Spreadsheet = False
      MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
      Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
      Stop
      Resume Next
    
  End Select
  
Resume ProcExit

End Function