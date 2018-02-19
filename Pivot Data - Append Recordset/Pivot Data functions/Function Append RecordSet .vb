Function FN_Append_Recordset_rsPUBLIC_Pivot(rsPivot As ADODB.Recordset, _
                                            sFieldName_RecordsetRowNumber As String, _
                                            Optional lRowNumber_Start As Long = 0, _
                                            Optional sFieldName_ExitOnNull As String = "") As Boolean

   'ADO objects
    Dim Fld As ADODB.Field
    Dim Cmd As ADODB.Command
    Dim Prm As ADODB.Parameter
  
   'Local Variables
    Dim i As Integer
    Dim lRowNumber As Long
    Dim lFieldCount As Long
    Dim bExit_DoLoop_LastRow_CMI_Pull As Boolean
   
   
   'Set function to TRUE
    FN_Append_Recordset_rsPUBLIC_Pivot = True
  
    On Error GoTo ProcErr
  
  
   'Instantiate public Recordset
    Set Cmd = New ADODB.Command
    Set Prm = New ADODB.Parameter


   '---------------------------------------------------------------------------------------------
   '  Append to PUBLIC recrodset rsPUBLIC_Pivot from recordset rsPivot
   '
   '  Append fields to the Public recordset: rsPUBLIC_Pivot
   '
   '  Field 1) ID field RowNumber
   '  Field 2) Add field as the FIRST field of the recordset
 
 
   '** Add field to recordset from rsPivot to rsPUBLIC_Pivot **
 
   If rsPUBLIC_Pivot.Fields.Count < 2 Then
   
     '1) Add Row Number  ID Field
     '   NOTE: This is the primary key field for this recordset as well as row number
      rsPUBLIC_Pivot.Fields.Append sFieldName_RecordsetRowNumber, adBigInt
      
     
     'Get field names
      For Each Fld In rsPivot.Fields
    
         'Append fields from recordset
          rsPUBLIC_Pivot.Fields.Append Fld.Name, Fld.Type, Fld.DefinedSize, Fld.Attributes
          
      Next
     
      
     '** Open Recordset rsPUBLIC_Pivot **
      rsPUBLIC_Pivot.Open
    
    End If
   
   
   'Move to first record of rsPivot
    rsPivot.MoveFirst
    
   'Set intial row count
    lRowNumber = lRowNumber_Start
   
   '------------ >>> Add records to rsPUBLIC_Pivot <<<-----------
    Do While Not rsPivot.EOF
  
      'Add New Record to PUBLIC recordset
        rsPUBLIC_Pivot.AddNew
  
      'Increment Field Value Key RowNumber
        lRowNumber = lRowNumber + 1
  
  
       'Get values from each field of the rsPivot passed to append the public recordset rsPUBLIC_Pivot
        For Each Fld In rsPUBLIC_Pivot.Fields
    
    
          'IF PivotTbl Key is Null then exit loop
          'NOTE: this will be the last row.
            If Fld.Name = sFieldName_ExitOnNull Then
        
                If IsNull(rsPivot(Fld.Name).Value) Then
   
                    bExit_DoLoop_LastRow_CMI_Pull = True
                    Exit For
                
                End If
            
            End If
 
           'Append KEY_FIELD_NAME value all other fields are appended after the else
            Select Case Fld.Name
        
              Case sFieldName_RecordsetRowNumber
              
                  Set Prm = Cmd.CreateParameter(, Fld.Type, adParamInput, Fld.DefinedSize, lRowNumber)
                  
              Case Else
              
                '--->> Create parameter and append to cmd object<<---
                'NOTE This ensure that the data coming from the Form is the right data type and right data size
                'For Example: If a value is copied into a cell that is read by the recordset does not require that cell to maintain the type of the copied value
                  Set Prm = Cmd.CreateParameter(, Fld.Type, adParamInput, Fld.DefinedSize, rsPivot(Fld.Name).Value)
            
            End Select
       

           '--->> Set Parameter Value to fields value <<---
            Fld.Value = Prm.Value  'rsPUBLIC_CMIworksheet_LineItems(Fld.Name).Value = Prm.Value

     Next
     
    'If Column BM value is NULL then stop populating recordset
     If bExit_DoLoop_LastRow_CMI_Pull = True Then
     
        Exit Do
     
     End If
     
    'Update all fields in the current records
     rsPUBLIC_Pivot.UpdateBatch
       
    'Move to next record for both recordsets
     rsPivot.MoveNext
    
  Loop
  

  
ProcExit:

'Close Recordset
 rsPivot.Close
 Set rsPivot = Nothing
 

 Exit Function

ProcErr:

  Select Case Err.Number

    Case 91, 424 'Object not found Note: This occurs on the rsTrackChanges close statement
      'Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
      Resume Next

    Case 3704 'Recordset is already closed
      Resume Next
      
    Case Else
      FN_Append_Recordset_rsPUBLIC_Pivot = False
      MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
      Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
      Stop
      Resume Next

  End Select

Resume ProcExit

End Function