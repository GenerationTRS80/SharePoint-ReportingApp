Public Function PubFN_sConn_SharePoint_PreSalesDb(Optional sTableName = "Employee_Opportunity")

    Dim sSharePointSite As String
    
5
     sSharePointSite = "https://thisURLtoConnectToSharePoint%presalesDB"
    

 'Read Only
    PubFN_sConn_SharePoint_PreSalesDb = "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=1;RetrieveIds=Yes;" & vbCrLf & _
                                        "DATABASE=" & sSharePointSite & ";" & vbCrLf & _
                                        "LIST=" & sTableName & ";"
                                        
End Function