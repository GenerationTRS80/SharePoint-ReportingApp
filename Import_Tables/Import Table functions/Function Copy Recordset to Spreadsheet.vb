Function SUB_SharePoint_CopyRecordset_to_Spreadsheet(xlWrkBk_ForecastTools As Excel.Workbook, _
                                                sTarget_WorksheetName As String, _
                                                rsSpeadsheet As ADODB.Recordset, _
                                                lStartRow As Long, _
                                                lStartColumn As Long, _
                                                Optional bAddHeader As Boolean = False, _
                                                Optional lClearData_Rows As Long = 1, _
                                                Optional lClearData_Columns As Long = 1, _
                                                Optional lSelectRow As Long = 1, _
                                                Optional lSelectColumn As Long = 1, _
                                                Optional sSet_NameRange_TargetWorksheet As String = "") As Boolean

'****************************************************************************
'*
'* NOTE: DO NOT USE MS EXCEL's  "Selection" for a substitute for a range of cells
'*          Excel will explicitly instantiate the "Selection" if you do use it
'*            You can 't close the instantiate selection
'*


'------- Local Variables ----------
 Dim lNumberFields As Long
 Dim lRows As Long
 Dim iCol As Integer
 Dim lHeaderRows As Long
 Dim lRowCount_RecordSet As Long
 
'------- Excel Objects -------
 Dim appXL As Excel.Application
 Dim xlWrkSht_TARGET As Excel.Worksheet
 Dim xlName As Name

'---------Range Objects--------
 Dim objCell_1 As Excel.Range
 Dim objCell_2 As Excel.Range
 Dim objRange As Excel.Range
 
 
'Set CopyRecordset to Spreadsheet to TRUE
 SUB_SharePoint_CopyRecordset_to_Spreadsheet = True
 
 
'---------Constants------------
 Const ADDITIONAL_FIELD_INCREMENT = 2 'This is required to get he last field in a recordset. Recordset field starts with ZEREO
 Const HEADER_ROWS_ADDED = 1
 
 On Error GoTo ProcErr

'Set target worksheet object
 Set xlWrkSht_TARGET = xlWrkBk_ForecastTools.Worksheets(sTarget_WorksheetName)
   
''Select worksheet Sheet1
' xlWrkSht_TARGET.Activate

    
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
        
'        'Copy field types to the first row of the worksheet
'         For iCol = lStartColumn To lNumberFields + lStartColumn - 1
'
'            xlWrkSht_TARGET.Cells(lStartRow + 1, iCol).Value = rsSpeadsheet.Fields(iCol - lStartColumn).Type
'
'         Next
'
'       'The number of rows that need to be added for the header
'        lHeaderRows = HEADER_ROWS_ADDED + 1
        
         lHeaderRows = HEADER_ROWS_ADDED
     Else
       
       'Zero rows are added when FALSE
        lHeaderRows = 0

 End If

'Copy Recordset to spreadsheet
 xlWrkSht_TARGET.Cells(lStartRow + lHeaderRows, lStartColumn).CopyFromRecordset rsSpeadsheet
 
'Recalculate worksheet
 xlWrkSht_TARGET.Calculate

'Set Select cell
 If lSelectRow <> 1 And lSelectColumn <> 1 Then
 
    Set objCell_1 = xlWrkSht_TARGET.Cells(lSelectRow, lSelectColumn)
    objCell_1.Select
    
 End If
 
'If worksheet  is visible then select cell G4
  If xlWrkSht_TARGET.Visible = xlSheetVisible Then

    xlWrkSht_TARGET.Range("G4").Select
 
 End If

 
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
                              
  End If
   

ProcExit:

'Close recordset
 rsSpeadsheet.Close
 Set rsSpeadsheet = Nothing


Exit Function


ProcErr:

  Select Case Err.Number
  
    Case 5 'Recordset error
    SUB_SharePoint_CopyRecordset_to_Spreadsheet = False
    MsgBox " The error # is " & Err.Number & vbCrLf & "Error with CopyRecordset Subroutine " & vbCrLf & "Send email to ITOCostModels with the error description", vbCritical + vbOKOnly
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume ProcExit
    
    Case 9 'Worksheet not found
    SUB_SharePoint_CopyRecordset_to_Spreadsheet = False
    MsgBox " The error # is " & Err.Number & vbCrLf & "Error with CopyRecordset Subroutine " & vbCrLf & "Send email to ITOCostModels with the error description", vbCritical + vbOKOnly
    Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
    Resume ProcExit
      
    Case 91, 424 'Hourglass Comand
      'Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
      Resume Next
      
    Case 1004
    '  CopyRecordset_to_Spreadsheet = False
    '  MsgBox " Too Many instances of Excel open. Close one or more instances", vbInformation + vbOKOnly
    '  Resume ProcExit
      Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
      Resume Next
      
    Case 3704 'Recordset is already closed
      Resume Next

    Case 3265 'Description of Item Can NOT be found in the recordset
    'If error then set SUB_SharePoint_CopyRecordset_to_Spreadsheet = False
      SUB_SharePoint_CopyRecordset_to_Spreadsheet = False
      MsgBox " The error # is " & Err.Number & vbCrLf & "Error with CopyRecordset Subroutine ", vbCritical + vbOKOnly
      Resume ProcExit

    Case -2147467259 'Steam Object can't be read because it is empty
    'If error then set SUB_SharePoint_CopyRecordset_to_Spreadsheet = False
      SUB_SharePoint_CopyRecordset_to_Spreadsheet = False
      MsgBox " The error # is " & Err.Number & vbCrLf & "Error with CopyRecordset Subroutine ", vbCritical + vbOKOnly
      Resume ProcExit
      
    Case Else
      SUB_SharePoint_CopyRecordset_to_Spreadsheet = False
      MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
      Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
      Stop
      Resume Next
    
  End Select

Resume ProcExit

End Function