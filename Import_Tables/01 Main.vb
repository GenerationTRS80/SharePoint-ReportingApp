Private Function aaMain(xlWrkSht_Button As Excel.Worksheet, Optional sTarget_WorksheetName As String)


 'Set workbook objects
  Dim xlWrkBk_Forecast As Excel.Workbook
  Dim rngCell_Table As Range
  Dim rngList_Tables As Range
  
 'Local Variables
  Dim bTestURL As Boolean
  Dim sTestURL_Test As String
     
  On Error GoTo ProcErr
  
'Turn Off Workbook Protection
  FN_Public_UnProtect_Workbook
  
 'Set default values
  PUBLIC_TEST_URL_YN = False
  
 'Instantiate Object
  Set xlWrkBk_Forecast = xlWrkSht_Button.Parent
  Set rngList_Tables = xlWrkBk_Forecast.Names("PDL_ListTables").RefersToRange
  
 'Set state of Excel application
  With Application
  
    .ScreenUpdating = False
    
  End With
  
  With xlWrkBk_Forecast.Worksheets("Main")
 
    .Range("B3").Select
    .Range("B3").Value = Now()
 
  End With
  
  ''    'NOTE: **** There is problem with "An error occurred in the secure channel support msxml3.dll" *****
  '
  ' '** Test if url in open logged inv**
  '  PUBLIC_TEST_URL_YN = TestURL(PUBLIC_URL_PRESALES)
  '
  '  If PUBLIC_TEST_URL_YN = False Then
  '
  '    MsgBox "Count not find/open Forecast Tool! at this url " & vbCrLf & PUBLIC_URL_PRESALES, vbExclamation + vbOKOnly, "Url Not Found"
  '    GoTo ProcExit
  '
  '  End If
  '
  ' 'Show url in Main page
  '  sTestURL_Test = "The URL " & PUBLIC_URL_PRESALES & vbCrLf & " The URL is " & PUBLIC_TEST_URL_YN


  '----------------------------------------------------------------------------------------------------------
  '
  '  ****** Run Import from SharePoint PreSales Db and load it into this spreadsheet *****

  'Write table in Sharepoint PreSales Database to Excel
  For Each rngCell_Table In rngList_Tables

    'Exit if no table name is returned from the name range PDL_ListTables
     If Len(rngCell_Table.Value) = 0 Then

       'Exit For loop
        Exit For

     End If

     'IMPORT SharePoint tables from the PreSales DB/Forecast Tool
      If abImport_SharepointLists_into_Spreadsheet(xlWrkSht_Button, rngCell_Table.Value) = False Then
      
       'If there is an error with the import then goto procexit
        GoTo ProcExit
        
      End If

  Next


 'Set cursor to Main page
 With xlWrkBk_Forecast.Worksheets("Main")
 
    .Activate
    .Range("C3").Select
    .Range("C3").Value = Now()
    .Range("B4").Value = sTestURL_Test
 
 End With


 Debug.Print ">>>> Programe has completed its run! Subroutine aaMain<<<<"
   
ProcExit:   'NOTE: ProcExit is the clean up function after error checking

 'Turn on screen updating
  With Application
  
    .ScreenUpdating = True
    
  End With

 'Turn On Workbook Protection
  FN_Public_Protect_Workbook


Exit Function

ProcErr:

  Select Case Err.Number
  
    Case 13  'Type Mismatch
      'MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbInformation + vbOKOnly, _"Function: Main Module: Forecat Tools Reports"
      Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
      Resume Next
    
    Case 91  'Object not found Note: This occurs on the rsTrackChanges close statement
      'Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
      Resume Next
      
    Case Else
      MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical
      Resume ProcExit
    
  End Select
  
Resume ProcExit

End Function