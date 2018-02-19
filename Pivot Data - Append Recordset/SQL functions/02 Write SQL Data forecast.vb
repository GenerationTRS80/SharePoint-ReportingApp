Private Function FN_Write_SQL_DataForecast(dStart_WeekDay As Date, _
                                lStart_WeekNumber As Long, _
                                iNumberWeeks_Ahead As Integer, _
                                sTblEmployee_Opp As String, _
                                sTblEmployee As String, _
                                sTblOpportunity As String, _
                                sTblClient As String) As String

    Dim sWeekYearNumber As String
    Dim dWeekDay As Date
    Dim sFormat_WeekDay As String
    Dim lYear As Long
    Dim lYearColumnAdjustment
    Dim sExcel_ColumnRef As String
    Dim sSQL As String
    Dim sSQL_Selected_HoursPerWeek
    
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
    
   'Create SQL statement columns for the number of weeks give by iNumberWeeks_Ahead
   'The SQL statment that represent each Week Hours is the sWeekYearNumber
    For i = 0 To iNumberWeeks_Ahead
    
        dWeekDay = DateAdd("d", 7 * i, dStart_WeekDay)
    
        lYear = Year(dStart_WeekDay)
        
        Debug.Print "lYear " & lYear
        'sWeekYearNumber = Format(i + lStart_WeekNumber, "00") & lYear
        
        sFormat_WeekDay = Format(dWeekDay, "mm/dd/yy")
          
        '*** Calculate the column to reference for weekly hours ***
        If lYear = 2016 Then
        
            lYearColumnAdjustment = -22
        
        Else
        
            'Make adjustment in columns for Year 2017
            'NOTE: In the series of Weekly hour columns there is a non Weekly hour column on the 21st column in  the series of column
            If i + lStart_WeekNumber < 21 Then
            
                lYearColumnAdjustment = 30
                
            Else
            
                lYearColumnAdjustment = 31
            
            End If
        
        End If

        
        sExcel_ColumnRef = "F" & i + lStart_WeekNumber + lYearColumnAdjustment
        'sSQL_Selected_HoursPerWeek = sSQL_Selected_HoursPerWeek & ", IIf(ISNULL(tblEmpOpp.[" & sExcel_ColumnRef & "]),0,Clng(Format(tblEmpOpp.[" & sExcel_ColumnRef & "],""0"")))  as [Hours per Week]" & vbCrLf
        sSQL_Selected_HoursPerWeek = sSQL_Selected_HoursPerWeek & ", IIf(ISNULL(tblEmpOpp.[" & sExcel_ColumnRef & "]),0,Clng(Format(tblEmpOpp.[" & sExcel_ColumnRef & "],""0""))) as [" & sFormat_WeekDay & "]" & vbCrLf


    Next i
    
    
   ' >>>>>>>>>>>>   Write SQL statement   <<<<<<<<<<<<

   sSQL = sSQL & "Select * From (Select "

'  'Key ,Description, Hours per Key Fields
   sSQL = sSQL & " tblEmpOpp.[F1] as [EmpOpp Key]" & vbCrLf
   
  'Employee Opportunity Description
   sSQL = sSQL & ", tblEmpOpp.[F2] as [Forecast Item Description]" & vbCrLf
   
'  'Table Keys for Employee and Opportunity table
'   sSQL = sSQL & ", tblEmpOpp.[F5] as [Employee ID]" & vbCrLf
'   sSQL = sSQL & ", tblEmpOpp.[F6] as [Opportunity ID]" & vbCrLf

  'Employee Fields
   sSQL = sSQL & ", tblEmp.[F2] AS [Full Name], tblEmp.[F3] AS [Last Name], tblEmp.[F5] AS [First Name]" & vbCrLf

  'Employee Team Fields
   sSQL = sSQL & ", tblEmpTeam.[F2] AS [Team Name], tblEmpTeam.[F7] AS [Team Sort]" & vbCrLf

  'Employee Skills
   sSQL = sSQL & ", tblSkill_Primary.[F2] as [Primary Skill]" & vbCrLf
   sSQL = sSQL & ", tblSkill_Secondary.[F2] as [Secondary Skill]" & vbCrLf
   
   sSQL = sSQL & ", tblTower.[F2] as [Tower]" & vbCrLf
   sSQL = sSQL & ", tblTower.[F6] as [Tower Sort]" & vbCrLf

  'Client
   sSQL = sSQL & ", tblClient.[F2] as [Client Name]" & vbCrLf
   
   
'  'Account group from Client table
'   sSQL = sSQL & ", tblClient.[F7] as [Client Account Group]" & vbCrLf

  'Opportunity Fields
   sSQL = sSQL & ", tblOpp.[F2] as [Nessie ID]" & vbCrLf
   sSQL = sSQL & ", tblOpp.[F3] as [Opp Name]" & vbCrLf

  'Opportunity Status
   sSQL = sSQL & ", tblStatus.[F2] as [Status]" & vbCrLf
   
'  'Request Type (the request type table need to be added for this field to function)
'   sSQL = sSQL & ", ""EMPTY"" as [Request Type]" & vbCrLf

  'Hour Category
   sSQL = sSQL & ", tblHourCat.[F2] as [Hour Category]" & vbCrLf
   sSQL = sSQL & ", tblHourCat.[F3] as [Hour Caterory Sort]" & vbCrLf

  'Solution Directors and  Bid Manager Opportunities
   sSQL = sSQL & ", tblEmp_SolutionDir.[F2] as [Solution Dir Opportunities List]" & vbCrLf
   sSQL = sSQL & ", tblEmp_BidMgr.[F2] as [Bid Manager Opportunities List]" & vbCrLf

'  'Opportunity Description
'   sSQL = sSQL & ", tblOpp.[F5] as [Opp Desc]" & vbCrLf

  'Opportunity ARR, TCV, Prob, Startegy, Solution, Offer
   sSQL = sSQL & ", tblOpp.[F6] as [Opp ARR], tblOpp.[F7] as [Opp TCV]" & vbCrLf
   sSQL = sSQL & ", tblOpp.[F8] as [Opp Terms], tblOpp.[F9] as [Opp Prob]" & vbCrLf
   sSQL = sSQL & ", tblOpp.[F10] as [Opp Strategy], tblOpp.[F11] as [Opp Solution]" & vbCrLf
   sSQL = sSQL & ", tblOpp.[F12] as [Opp Offer Review], tblOpp.[F13] as [Opp Orals]" & vbCrLf
   sSQL = sSQL & ", tblOpp.[F14] as [Opp Downselect], tblOpp.[F15] as [Opp Due Diligence]" & vbCrLf
   sSQL = sSQL & ", tblOpp.[F16] as [Opp BAFO], tblOpp.[F17] as [Opp Negotiation]" & vbCrLf
   sSQL = sSQL & ", tblOpp.[F18] as [Opp Close], tblOpp.[F19] as [Opp Handover]" & vbCrLf

  'Opportunity Comment Fields
   sSQL = sSQL & ", tblOpp.[F28] as [Opp Issues Risks]" & vbCrLf
   sSQL = sSQL & ", tblOpp.[F29] as [Opp Status NextStep]" & vbCrLf

  'Opportunity Table Modified Date
  'NOTE Opportunity Last Updated By is the Full Name field from the employee table joined by SharePoint Editor
   sSQL = sSQL & ", tblEmp_SharePointID.[F2] as [Opportunity Updated By]" & vbCrLf
   sSQL = sSQL & ", tblOpp.[F43] as [Opportunity Modified Date]" & vbCrLf


  ' >>>> Column Hours Per Week <<<<
  'NOTE: This string is calculated sSQL_Selected_HoursPerWeek is constructed above
   sSQL = sSQL & sSQL_Selected_HoursPerWeek


  'Employee Opportunity Set Filter
   sSQL = sSQL & ", tblEmpOpp.[F7] as [Filter]" & vbCrLf

  'Report Filter by  Scott Archer, David Lancaster,Jennifer Shea
   sSQL = sSQL & ", tblEmp.[F12] AS [David L Report Filter]" & vbCrLf
   sSQL = sSQL & ", tblEmp.[F13] AS [Jennifer Shea Report Filter]" & vbCrLf
   sSQL = sSQL & ", tblEmp.[F14] AS [Scott Archer Report Filter]" & vbCrLf

  'Employee Opportunity date fields
   sSQL = sSQL & ", tblEmpOpp.[F86] as [Forecast Modified Date]" & vbCrLf
   sSQL = sSQL & ", tblEmpOpp.[F87] as [Forecast Created Date]" & vbCrLf


  'Calculated the last modified time Date fields
  'NOTE: There is an archive date (tblEmpOpp.[F77])that needs to be compared to Share Point Modified date
  '      to determine the TRUE/ACTUAL modified date.
  '      If the Archive Date [F77] is null then use [SharePoint Modified Date]
  '      If the Archive Date [F77] is NOT Null then check to see if the [SharePoint Modified Date] is < ' ' before Feb 2 2017
  '      If the [SharePoint modified date] is BEFORE Feb 2 2017 then use the Archive Date [F77]
  '      If the [SharePoint modified date] is AFTER Feb 2 2017 then use the [SharePoint modified date]

   sSQL = sSQL & ", IIf(IIf(ISNULL(tblEmpOpp.[F77]),FALSE,TRUE)" & vbCrLf
  'Original Date: sSQL = sSQL & ", IIf([Forecast Modified Date]<DateSerial(2017,6,16)" & vbCrLf
   sSQL = sSQL & ", IIf([Forecast Modified Date]<DateSerial(2017,6,21)+15/24" & vbCrLf
   sSQL = sSQL & ", CDate(Format(tblEmpOpp.[F77],""mm/dd/yy"")),CDate(Format([Forecast Modified Date],""mm/dd/yy"")))"
   sSQL = sSQL & ", CDate(Format([Forecast Modified Date],""mm/dd/yy""))) as [Forecast Update]" & vbCrLf
   
  '  >>> Organizational Unit (Last Field) <<<
   sSQL = sSQL & ", tblOrgUnit.[F2] AS [Organization Unit]"
   
  'Forecast Line Item Removed and Removed Date
   sSQL = sSQL & ", tblEmpOpp.[F78] as [Remove Forecast]" & vbCrLf
   sSQL = sSQL & ", Format(tblEmpOpp.[F79],""mm/dd/yy"") as [Removed Forecast Update]" & vbCrLf
    

  '-------------------------------------------------------------------------------------------------------------
  '
  '    ** FROM CLAUSE **
  
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
   'sSQL = sSQL & ") as tblPivotHours" & vbCrLf


   '------------------------- *** Where Clause *** ----------------------------------
   '
   '    NOTE:  1) Employee Organization Unit equal PreSales or the value 2
   '           2) Exclude the first row from table tblEmployee_Opportunity
   
    sSQL = sSQL & " Where tblEmp.[F18]=2 and tblEmpOpp.[F1]>1) as tblPivotHours"

    'sSQL = sSQL & " Where tblEmpOpp.[F1]>1) as tblPivotHours"
 
    'sSQL = sSQL & " WHERE tblPivotHours.[Hour Category]= ""Deal Hours""  ;"
    'sSQL = sSQL & " WHERE tblPivotHours.[Status]= ""DEAD"" ;"
    'sSQL = sSQL & " WHERE tblPivotHours.[OppID]= 29 ;"
 
    sSQL = sSQL & " ORDER By tblPivotHours.[Forecast Update] desc, tblPivotHours.[Full Name], tblPivotHours.[Hour Caterory Sort] "
    sSQL = sSQL & ",tblPivotHours.[Status] ,tblPivotHours.[Client Name] , tblPivotHours.[Opp Name]"

    
    FN_Write_SQL_DataForecast = sSQL

End Function