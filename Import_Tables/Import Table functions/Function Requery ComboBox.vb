Public Sub SUB_Requery_ComboBox(cboComboBox As ComboBox, Optional sListFillRange)
'*
'* NOTE Combobox should have all the same properties and the ActiveX OLE ComboBox See One page notes
'*

   'Local Variables
    Dim sSelectedText As String

   'Turn on Error Handeling
    On Error GoTo ProcErr
        
   'Requery combo box
    With cboComboBox
        
        .ListFillRange = sListFillRange
        '.ShowDropButtonWhen = fmShowDropButtonWhenAlways

    End With
                                                                   
   'Set combo box to the first value of the FileNameReference for the tower
    sSelectedText = cboComboBox.List(0, 1)
    cboComboBox.Text = sSelectedText
    
ProcExit:
 

ProcErr:

  Select Case Err.Number

    Case 13 'Cancel button hit on input box
      Resume ProcExit
      
    Case 94 'Invalid use of Null FileNameReference not found (see NONE or Tower has not had a FileNameReference set to it see PullDown List sheet)
      Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & "The source " & Err.Source & vbCrLf & _
      "Tower has not had a FileNameReference set to it in the PullDown List sheet"
      
      Resume Next
      
    Case 3704 'Operation to allowed with Object is closed
      Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & "The source " & Err.Source
      Resume Next

    Case -2147467259 'Operation to allowed with Object is closed
      Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & "The source " & Err.Source
      Resume ProcExit
      
    Case Else 'IMdcCombo failed
    
     '** This will capture any error and show its number and description
      Debug.Print " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & "The source " & Err.Source
      MsgBox " The error # is " & Err.Number & vbCrLf & "Description " & Err.Description & vbCrLf & vbCrLf & " The source " & Err.Source, vbCritical
      Stop
      Resume Next

  End Select

Resume ProcExit

End Sub