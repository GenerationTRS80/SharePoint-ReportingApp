Private Function FN_Write_SQL_IssueRisks(sTblEmployee_Opp As String, _
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
    
 
   ' >>>>>>>>>>>>   Write SQL statement   <<<<<<<<<<<<

   sSQL = sSQL & "Select * From (Select "

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

  ''Opportunity Description
  ' sSQL = sSQL & ", tblOpp.[F5] as [Opp Desc]" & vbCrLf

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
  'sSQL = sSQL & ") as tblPivotHours" & vbCrLf


  'Where Clause
  'NOTE: tblEmp.[18]=2 is equal Employee Org= PRE-Sales
   sSQL = sSQL & " Where tblEmp.[F18]=2) as tblPivotHours"
  
  'sSQL = sSQL & " WHERE tblPivotHours.[Hour Category]= ""Deal Hours""  ;"
  'sSQL = sSQL & " WHERE tblPivotHours.[Status]= ""DEAD"" ;"
  'sSQL = sSQL & " WHERE tblPivotHours.[OppID]= 29 ;"
 
   sSQL = sSQL & " ORDER By  tblPivotHours.[Full Name], tblPivotHours.[Hour Caterory Sort] "
   sSQL = sSQL & ",tblPivotHours.[Status] ,tblPivotHours.[Client Name] , tblPivotHours.[Opp Name]"

    
   FN_Write_SQL_IssueRisks = sSQL

End Function