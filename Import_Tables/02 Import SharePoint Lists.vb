Private Function abImport_SharepointLists_into_Spreadsheet(xlWrkSht_Button As Excel.Worksheet, _
                                                        Optional sTableName01 As String) As Boolean


'ADO Objects
  Dim rsSharePoint_PreSales As ADODB.Recordset
  Dim Cmd As ADODB.Command
  Dim Conn As ADODB.Connection
  Dim Rec As ADODB.Record

  
'Local Variables
  Dim sSQL As String
  Dim sConnString As String
  Dim sListFillRange_Address As String
  Dim sNameRange As String
  Dim lRowCount As Long
  Dim i As Integer
  
  Dim sUserID As String
  Dim sPassWord As String
      
  On Error GoTo ProcErr
  
'Set default value for function
  abImport_SharepointLists_into_Spreadsheet = False
  
'Instantiate objects
  Set rsSharePoint_PreSales = New ADODB.Recordset
  Set Cmd = New ADODB.Command
  Set Conn = New ADODB.Connection
    
'Open Connection
  sConnString = PubFN_sConn_SharePoint_PreSalesDb(sTableName01)
  
 'Write SQL statement
  If sTableName01 = "Employee_Opportunity" Then
  
    sSQL = "Select distinct * from " & sTableName01 & " Order by SharePointModifiedDate desc"

  Else
  
    sSQL = "Select distinct * from " & sTableName01
  
  End If
  
    
'*** Using command Object to get recordset ***

 'Open Connection
   Conn.Open sConnString


'Set command object
  With Cmd

    .ActiveConnection = Conn
    .CommandText = sSQL
    .CommandType = adCmdText
    .CommandTimeout = 60  'NOTE: if user isn't logged into PreSale DB ie. have Forecast Tool website open. Then app will timeout
    
  End With

 'Set recordset to returned value from Command Object
  With rsSharePoint_PreSales
  
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    
  End With
  
' *** Execute command object ***
  Set rsSharePoint_PreSales = Cmd.Execute

'Count the number of rows
  Do While Not rsSharePoint_PreSales.EOF
  
        lRowCount = lRowCount + 1
        rsSharePoint_PreSales.MoveNext
  Loop
  

  With rsSharePoint_PreSales
  
            .MoveFirst

  End With
  
 
'Copy Recordset to Spreadheet
 sNameRange = "tbl" & sTableName01
 
 If SUB_SharePoint_CopyRecordset_to_Spreadsheet(xlWrkSht_Button.Parent, _
                                                sTableName01, _
                                                rsSharePoint_PreSales, _
                                                3, 1, _
                                                True, _
                                                3000, 50, _
                                                1, 1, _
                                                sNameRange) = False Then
 
    Debug.Print "**** Error Exited CopyRecordset Sub ******"
    GoTo ProcExit

 End If

'Set function to TRUE when all actions are completed
 abImport_SharepointLists_into_Spreadsheet = True
 
ProcExit:
  
 'Close Connection object
  Conn.Close
  Set Conn = Nothing
  
  rsSharePoint_PreSales.Close
  Set rsSharePoint_PreSales = Nothing
   
  Debug.Print ">>>> Programe has completed its run! <<<<"
  
Exit Function


ProcErr:

  Select Case Err.Number
  
    Case 91, 424  'Object not found Note: This occurs on the rsTrackChanges close statement
      'Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
      Resume Next
    
    Case 3704 'Recordset empty End program to stop more errors
      Resume Next
      
    Case -2147217865 'Can NOT find the object
        MsgBox "Can not find table " & sTableName01 & " This table was not updated!" _
      & vbCrLf & vbCrLf & "Please, email the ITOPursuitSite mailbox with a description of this problem listed below!," _
      & vbCrLf & vbExclamation + vbOKOnly, "Function: Import SharepointLists into Spreadsheet Module: Forecast Tool Report"
      Resume ProcExit
      
    Case -2147467259 'Could not find installable ISAM
      Debug.Print "The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
      MsgBox "MS Access is not installed on your computer" _
      & vbCrLf & vbCrLf & "Go to the Start Menu and select All Programs" _
      & vbCrLf & "then go to Office Applications and Click on Microsoft Access 2010" _
      & vbCrLf & vbCrLf & "If you are unable to install MS Access then send an email to itopursuitsites@atos.net mailbox" _
      , vbExclamation + vbOKOnly, "Function: Import SharepointLists into Spreadsheet Module: Forecast Tool Report"
      Resume ProcExit

    Case Else
      MsgBox "Description " & Err.Description & vbCrLf & "The error # is " & Err.Number & vbCrLf & "The source " & Err.Source, vbCritical
      Resume ProcExit
    
  End Select

Resume ProcExit

End Function