Private Function FN_Write_SQL_Pivot(dWeekDay As Date, _
                                lWeekNumber As Long, _
                                sFieldName_for_KeyField As String, _
                                sTblEmployee_Opp As String, _
                                sTblEmployee As String, _
                                sTblOpportunity As String, _
                                sTblClient As String) As String


    Dim sWeekYearNumber As String
    Dim lYear As Long
    Dim lYearColumnAdjustment
    Dim sExcel_ColumnRef As String
    Dim sCalc_Hours As String
    
   'Table Vars
    Dim sTblEmployee_SDirector As String
    Dim sTblEmployee_BidMgr As String
    Dim sTblEmployeeTeam As String
    Dim sTblHourCategory As String
    Dim sTblStatus As String
    

    sTblEmployee_SDirector = sTblEmployee  'Solution Director Tables
    sTblEmployee_BidMgr = sTblEmployee 'Bid Manager table based off  employee table
    sTblEmployee_SharePoint_ID = sTblEmployee
    sTblEmployeeTeam = "[Employee_Team$A3:V13]"
    sTblHourCategory = "[Hour_Category$A3:C20]"
    sTblStatus = "[New_Status$A3:B13]"
    sTblSkill_CategoryPrimary = "[Skill_Category$A3:N103]"
    sTblSkill_CategorySecondary = "[Skill_Category$A3:N103]"
    sTblTower = "[Tower$A3:M100]"
    sTblEmployee_OrgUnit = "[Employee_Org_Unit$A3:H15]"
    

   'Calculate Week Number
    lYear = Year(dWeekDay)
    sWeekYearNumber = Format(lWeekNumber, "00") & lYear
              
              
   '*** Calculate the column to reference for weekly hours ***
    If lYear = 2016 Then
      
        lYearColumnAdjustment = -22
      
     Else
      
        'Make adjustment in columns for Year 2017
        'NOTE: In the series of Weekly hour columns there is a non Weekly hour column on the 21st column in the series of column
         If lWeekNumber < 21 Then
          
            lYearColumnAdjustment = 30
              
         Else
          
            lYearColumnAdjustment = 31
          
         End If
      
     End If
      
    'Create SQL column reference
     sExcel_ColumnRef = "F" & lWeekNumber + lYearColumnAdjustment
        
        
    ' >>>>>>>>>>>>   Write SQL statement   <<<<<<<<<<<<

     sSQL = sSQL & "Select * From ("
     sSQL = sSQL & "Select #" & dWeekDay & "# as [Week Date]" & vbCrLf

    'Key ,Description, Hours per Key Fields
     sSQL = sSQL & ", tblEmpOpp.[F1] as [" & sFieldName_for_KeyField & "]" & vbCrLf
     sSQL = sSQL & ", tblEmpOpp.[F2] as [Employee_Opp Desc]" & vbCrLf
     sSQL = sSQL & ", tblEmpOpp.[F7] as [Forecast Filter]" & vbCrLf
     
    '--------------------------------------------------------------------------------------
    '** February 27th column data can't function with an greater than or less than sign
     If sExcel_ColumnRef = "F39" Then
     
        sCalc_Hours = "IIf(ISNULL(tblEmpOpp.[" & sExcel_ColumnRef & "]),0,Clng(Format(tblEmpOpp.[" & sExcel_ColumnRef & "],""0"")))"
        'sCalc_Hours = 0
      
     Else
     
        sCalc_Hours = "IIf(ISNULL(tblEmpOpp.[" & sExcel_ColumnRef & "]),0,Clng(Format(tblEmpOpp.[" & sExcel_ColumnRef & "],""0"")))"
       ' sSQL = sSQL & ", IIf(ISNULL(tblEmpOpp.[" & sExcel_ColumnRef & "]),0,Clng(Format(tblEmpOpp.[" & sExcel_ColumnRef & "],""0"")))  as [Hours per Week]" & vbCrLf
     
     End If
     
     
     sSQL = sSQL & ", " & sCalc_Hours & " as [Hours per Week]" & vbCrLf

    'Employee Fields
     sSQL = sSQL & ", tblEmp.[F2] AS [Full Name], tblEmp.[F3] AS [Last Name], tblEmp.[F5] AS [First Name]" & vbCrLf
   
    'Employee Skills
     sSQL = sSQL & ", tblSkill_Primary.[F2] as [Primary Skill]" & vbCrLf
     sSQL = sSQL & ", tblSkill_Secondary.[F2] as [Secondary Skill]" & vbCrLf
     sSQL = sSQL & ", tblTower.[F2] as [Tower]" & vbCrLf
  
    'Employee Team Fields
     sSQL = sSQL & ", tblEmpTeam.[F2] AS [Team Name], tblEmpTeam.[F7] AS [Team Sort]" & vbCrLf
   
    'Report Filter by  Scott Archer, David Lancaster,Jennifer Shea
     sSQL = sSQL & ", tblEmp.[F12] AS [David L Report Filter]" & vbCrLf
     sSQL = sSQL & ", tblEmp.[F13] AS [Jennifer Shea Report Filter]" & vbCrLf
     sSQL = sSQL & ", tblEmp.[F14] AS [Scott Archer Report Filter]" & vbCrLf

    'Opportunity Status
     sSQL = sSQL & ", tblStatus.[F2] as [Status]" & vbCrLf

    'Opportunity Fields
     sSQL = sSQL & ", tblOpp.[F2] as [Nessie ID]" & vbCrLf
     sSQL = sSQL & ", tblOpp.[F3] as [Opp Name]" & vbCrLf
     sSQL = sSQL & ", tblOpp.[F5] as [Opp Desc]" & vbCrLf

    'Client
     sSQL = sSQL & ", tblClient.[F2] as [Client Name]" & vbCrLf
     sSQL = sSQL & ", tblClient.[F7] as [Client Account Group]" & vbCrLf
    
    'Solution Directors and  Bid Manager Opportunities
     sSQL = sSQL & ", tblEmp_SolutionDir.[F2] as [Solution Dir Opportunities List]" & vbCrLf
     sSQL = sSQL & ", tblEmp_BidMgr.[F2] as [Bid Manager Opportunities List]" & vbCrLf

    'Hour Category
     sSQL = sSQL & ", tblHourCat.[F2] as [Hour Category]" & vbCrLf
     sSQL = sSQL & ", tblHourCat.[F3] as [Hour Caterory Sort]" & vbCrLf

    'Opportunity Comment Fields
     sSQL = sSQL & ", tblOpp.[F28] as [Opp Issues Risks]" & vbCrLf
     sSQL = sSQL & ", tblOpp.[F29] as [Opp Status NextStep]" & vbCrLf
   
  
    'Employee Opportunity date fields
     sSQL = sSQL & ", tblEmpOpp.[F86] as [Forecast Modified Date]" & vbCrLf
     sSQL = sSQL & ", tblEmpOpp.[F87] as [Forecast Created Date]" & vbCrLf
     
    'Opportunity Table Modified Date
    'NOTE Opportunity Last Updated By is the Full Name field from the employee table joined by SharePoint Editor
     sSQL = sSQL & ", tblEmp_SharePointID.[F2] as [Opportunity Updated By]" & vbCrLf
     sSQL = sSQL & ", tblOpp.[F43] as [Opportunity Modified Date]" & vbCrLf
   
    'Calculated the last modified time Date fields
    'NOTE: There is an archive date (tblEmpOpp.[F77])that needs to be compared to Share Point Modified date
    '      to determine the TRUE/ACTUAL modified date.
    '      If the Archive Date [F77] is null then use [SharePoint Modified Date]
    '      If the Archive Date [F77] is NOT Null then check to see if the [SharePoint Modified Date] is < before Feb 2 2017
    '      If the [SharePoint modified date] is BEFORE Feb 2 2017 then use the Archive Date [F77]
    '      If the [SharePoint modified date] is AFTER Feb 2 2017 then use the [SharePoint modified date]

     sSQL = sSQL & ", IIf(IIf(ISNULL(tblEmpOpp.[F77]),FALSE,TRUE)" & vbCrLf
     sSQL = sSQL & ", IIf([Forecast Modified Date]<DateSerial(2017,6,21)+15/24" & vbCrLf
     sSQL = sSQL & ", CDate(Format(tblEmpOpp.[F77],""mm/dd/yy"")),CDate(Format([Forecast Modified Date],""mm/dd/yy"")))"
     sSQL = sSQL & ", CDate(Format([Forecast Modified Date],""mm/dd/yy""))) as [Forecast Update]" & vbCrLf
   
    ' >>> Organizational Unit (Last Field) <<<
     sSQL = sSQL & ", tblOrgUnit.[F2] AS [Organization Unit]"
    
    'Forecast Line Item Removed and Removed Date
     sSQL = sSQL & ", tblEmpOpp.[F78] as [Removed Forecast]" & vbCrLf
     sSQL = sSQL & ", tblEmpOpp.[F79] as [Removed Forecast Update]" & vbCrLf


    '*From Clause
     sSQL = sSQL & " From " & sTblEmployee_OrgUnit & " as tblOrgUnit" & vbCrLf

     sSQL = sSQL & " INNER JOIN (" & sTblEmployee_SharePoint_ID & " as tblEmp_SharePointID" & vbCrLf
     sSQL = sSQL & " INNER JOIN (" & sTblTower & " as tblTower" & vbCrLf
     sSQL = sSQL & " INNER JOIN (" & sTblSkill_CategorySecondary & " as tblSkill_Secondary" & vbCrLf
     sSQL = sSQL & " INNER JOIN (" & sTblSkill_CategoryPrimary & " as tblSkill_Primary" & vbCrLf
     sSQL = sSQL & " INNER JOIN (" & sTblEmployee_BidMgr & " as tblEmp_BidMgr" & vbCrLf
     sSQL = sSQL & " INNER JOIN (" & sTblEmployee_SDirector & " as tblEmp_SolutionDir" & vbCrLf
     sSQL = sSQL & " INNER JOIN (" & sTblStatus & " as tblStatus" & vbCrLf
     sSQL = sSQL & " INNER JOIN (" & sTblClient & " as tblClient" & vbCrLf
     sSQL = sSQL & " INNER JOIN (" & sTblHourCategory & " as tblHourCat" & vbCrLf
     sSQL = sSQL & " INNER JOIN (" & sTblOpportunity & " as tblOpp" & vbCrLf
     sSQL = sSQL & " INNER JOIN (" & sTblEmployeeTeam & " as tblEmpTeam" & vbCrLf
     sSQL = sSQL & " INNER JOIN (" & sTblEmployee & " as tblEmp" & vbCrLf

     sSQL = sSQL & " INNER JOIN " & sTblEmployee_Opp & " as tblEmpOpp" & vbCrLf
     sSQL = sSQL & " ON tblEmp.[F1]=tblEmpOpp.[F5])" & vbCrLf
     sSQL = sSQL & " ON tblEmpTeam.[F1]=tblEmp.[F9])" & vbCrLf
     sSQL = sSQL & " ON tblOpp.[F1]=tblEmpOpp.[F6])" & vbCrLf
     sSQL = sSQL & " ON tblHourCat.[F1]=tblOpp.[F35])" & vbCrLf
     sSQL = sSQL & " ON tblClient.[F1]=tblOpp.[F21])" & vbCrLf
     sSQL = sSQL & " ON tblStatus.[F1]=tblOpp.[F23])" & vbCrLf
     sSQL = sSQL & " ON tblEmp_SolutionDir.[F1]=tblOpp.[F22])" & vbCrLf
     sSQL = sSQL & " ON tblEmp_BidMgr.[F1]=tblOpp.[F25])" & vbCrLf
     sSQL = sSQL & " ON tblSkill_Primary.[F1]=tblEmp.[F15])" & vbCrLf
     sSQL = sSQL & " ON tblSkill_Secondary.[F1]=tblEmp.[F16])" & vbCrLf
     sSQL = sSQL & " ON tblTower.[F1]=tblEmp.[F17])" & vbCrLf
     sSQL = sSQL & " ON tblEmp_SharePointID.[F19]=tblOpp.[F41])" & vbCrLf

     sSQL = sSQL & " ON tblOrgUnit.[F1]=tblEmp.[F18]" & vbCrLf
    ' sSQL = sSQL & ") as tblPivotHours" & vbCrLf

    '------------------------- *** Where Clause *** ----------------------------------
    '
    '    NOTE:  1) Filter for hours greater than 0 2) Employee Organization Unit equal PreSales or the value 2
    '           3) Exclude the first row from table tblEmployee_Opportunity
    

    '--------------------------------------------------------------------------------------
    '** February 27th column data can't function with an greater than or less than sign
     If sExcel_ColumnRef = "F39" Then
    
      sSQL = sSQL & " Where tblEmpOpp.[" & sExcel_ColumnRef & "] Is Not Null and tblEmp.[F18]=2 ) as tblPivotHours"
       
     Else
    
      sSQL = sSQL & " Where " & sCalc_Hours & ">0 and tblEmp.[F18]=2) as tblPivotHours"
      
     End If
    
    
    '  sSQL = sSQL & " Where tblEmpOpp.[" & sExcel_ColumnRef & "]>0 and tblEmp.[F18]=2) as tblPivotHours"
    '  sSQL = sSQL & " Where tblEmpOpp.[" & sExcel_ColumnRef & "] Is Not Null ) as tblPivotHours"
    '  sSQL = sSQL & " WHERE CDbl(tblPivotHours.[Hours per Week])>0 ;"
    '  sSQL = sSQL & " WHERE tblPivotHours.[Hour Category]= ""Deal Hours""  ;"
    '  sSQL = sSQL & " WHERE tblPivotHours.[Status]= ""DEAD"" ;"
    '  sSQL = sSQL & " WHERE tblPivotHours.[OppID]= 29 ;"


    '------------------------- *** Order By clause *** ----------------------------------
    '
    '
    
     sSQL = sSQL & " ORDER By tblPivotHours.[Week Date], tblPivotHours.[Full Name], tblPivotHours.[Hour Caterory Sort] "
     sSQL = sSQL & ",tblPivotHours.[Status] ,tblPivotHours.[Client Name] , tblPivotHours.[Opp Name]"


     FN_Write_SQL_Pivot = sSQL


End Function