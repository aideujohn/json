# json
for caresharing



Private JiraService As New MSXML2.XMLHTTP60
Public ResponseLength As Long 'contains the length of the last response string

Public Username As String 'must fill to use the class
Public Password As String 'must fill to use the class

Public IntranetLink As String 'must fill to use the class, local jira link e.g "http://eudca-jira01/jira/rest/api/2/"
Public Issuekey As String 'must fill to use the update/get function, have to adress the ticket to update/get e.g "TESTSYS-157"
Public Projectkey As String 'as pivot variable, will hold project name
Public ParentIssue As String 'as pivot variable, will parent issue key

Public Response As String 'Will contain the returning response of a Rest request
Public Status As String 'Will contain the returning status of a rest request

Public SendData As String 'Will contain the information to be sent to the rest api to perform a command


Public Function CreateTicket()
With JiraService
        .Open "POST", IntranetLink & "issue", False  'Jira Link
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Accept", "application/json"
        .setRequestHeader "Authorization", "Basic " & EncodeBase64(Username & ":" & Password)
        On Error Resume Next
        .send SendData
        Response = .responseText
        Status = .Status & " | " & .statusText
        ResponseLength = Len(Response)
        
    End With
End Function
Public Function UpdateTicket()
With JiraService
        .Open "PUT", IntranetLink & "issue/" & Issuekey, False  'Jira Link
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Accept", "application/json"
        .setRequestHeader "Authorization", "Basic " & EncodeBase64(Username & ":" & Password)
        On Error Resume Next
        .send SendData
        Response = .responseText
        Status = .Status & " | " & .statusText
        ResponseLength = Len(Response)
        
    End With
End Function
Public Function GetTicket()
With JiraService
        .Open "GET", IntranetLink & "search?jql=issuekey=" & Issuekey, False   'Jira Link
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Accept", "application/json"
        .setRequestHeader "Authorization", "Basic " & EncodeBase64(Username & ":" & Password)
        On Error Resume Next
        .send
        Response = .responseText
        Status = .Status & " | " & .statusText
        'Sheets("SysProject Planning").Cells(70, 5).Value = Response
        ResponseLength = Len(Response)
    End With
End Function


Private Function EncodeBase64(text As String) As String

    Dim arrData() As Byte
    arrData = StrConv(text, vbFromUnicode)
  
    Dim objXML As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement
  
    Set objXML = New MSXML2.DOMDocument
    Set objNode = objXML.createElement("b64")
  
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.text
  
    Set objNode = Nothing
    Set objXML = Nothing
    
End Function

'ONLY WORKS FOR SUBTASK, have to remove a few lines to make it work with main tickets
'Must search a single SUBTASK without search parameters, this one gets the value on a field with a single value (no brackets with multiple fields inside),
'Won't work with fields with many values such as status or issuetype
Public Function FieldSeekerSingleValue(Jresponse As String, Fieldname As String) As String

Dim Pointer As Long
Dim Length As Long
Dim FieldisNull As Boolean

Pointer = InStr(Jresponse, "{") 'Take out first expand bracket
Jresponse = Right(Jresponse, ResponseLength - Pointer)

Pointer = InStr(Jresponse, "{") + Pointer 'Take out issue bracket (this function only works for single issue searches)
Jresponse = Right(Jresponse, ResponseLength - Pointer)

Pointer = InStr(Jresponse, "{") + Pointer 'Take out fields bracket
Jresponse = Right(Jresponse, ResponseLength - Pointer)

Pointer = InStr(Jresponse, """parent""") + Pointer - 1 'remove to work with main task ********
Jresponse = Right(Jresponse, ResponseLength - Pointer)

Pointer = InStr(Jresponse, "}") + Pointer 'remove to work with main task *********
Jresponse = Right(Jresponse, ResponseLength - Pointer)

Pointer = InStr(Jresponse, Fieldname) + Pointer - 1 'We keep the field
Jresponse = Right(Jresponse, ResponseLength - Pointer)

Pointer = InStr(Jresponse, ":") + Pointer 'we take out the field name
Jresponse = Right(Jresponse, ResponseLength - Pointer)
If InStr(Jresponse, "null") = 1 Or InStr(Jresponse, "null") = 2 Then FieldisNull = True

On Error GoTo NoComa

Pointer = InStr(Jresponse, ",") 'We keep only the field value
InCaseofError:
Jresponse = Left(Jresponse, Pointer - 1)

Length = Len(Jresponse)
If Not IsNumeric(Jresponse) Then
Jresponse = Left(Jresponse, Length - 1)
Jresponse = Right(Jresponse, Length - 2)
End If
If FieldisNull = True Then Jresponse = ""
FieldSeekerSingleValue = Jresponse

Exit Function

NoComa:
Pointer = InStr(Jresponse, "}")
GoTo InCaseofError

End Function

'ONLY WORKS FOR SUBTASK, have to remove a few lines to make it work with main tickets
'Must search a single SUBTASK without search parameters, this one gets the value on a field with a single value (no brackets with multiple fields inside),
'Intended for fields with multiple values inside such as issuetype or project
'fieldname por Project would be """project""" and then as subfield name there would be """name""" , """id""", etc
Public Function FieldSeekerMultiValue(Jresponse As String, Fieldname As String, SubFieldName) As String

Dim FieldisNull As Boolean
Dim Pointer As Long
Dim MultiValue As Boolean
Dim MultiValueJresponse As String
Dim MultiValueJresponsePivot As String
Dim MultiValuePointer As Long
Dim Length As Long

Pointer = InStr(Jresponse, "{") 'Take out first expand bracket
Jresponse = Right(Jresponse, ResponseLength - Pointer)

Pointer = InStr(Jresponse, "{") + Pointer 'Take out issue bracket (this function only works for single issue searches)
Jresponse = Right(Jresponse, ResponseLength - Pointer)

Pointer = InStr(Jresponse, "{") + Pointer 'Take out fields bracket
Jresponse = Right(Jresponse, ResponseLength - Pointer)

Pointer = InStr(Jresponse, """parent""") + Pointer - 1 'remove to work with main task ********
Jresponse = Right(Jresponse, ResponseLength - Pointer)

Pointer = InStr(Jresponse, "}") + Pointer 'remove to work with main task *********
Jresponse = Right(Jresponse, ResponseLength - Pointer)

Pointer = InStr(Jresponse, Fieldname) + Pointer - 1 'We keep the field
Jresponse = Right(Jresponse, ResponseLength - Pointer)



Pointer = InStr(Jresponse, ":") + Pointer 'we take out the field name
Jresponse = Right(Jresponse, ResponseLength - Pointer)

If InStr(Jresponse, "null") = 1 Or InStr(Jresponse, "null") = 2 Then FieldisNull = True

'****************************
If InStr(Jresponse, "[") = 1 Then
    MultiValue = True
Else:
    MultiValue = False
End If
'****************************

Pointer = InStr(Jresponse, SubFieldName) + Pointer - 1 'We keep the field
Jresponse = Right(Jresponse, ResponseLength - Pointer)

Pointer = InStr(Jresponse, ":") + Pointer 'we take out the field name
Jresponse = Right(Jresponse, ResponseLength - Pointer)

'****************************
If MultiValue = True Then
    MultiValueJresponse = Jresponse
    MultiValuePointer = Pointer
    MultiValuePointer = InStr(MultiValueJresponse, "]")
    MultiValueJresponse = Left(MultiValueJresponse, MultiValuePointer - 1)
End If
'****************************

On Error GoTo NoComa

Pointer = InStr(Jresponse, ",") 'We keep only the field value
InCaseofError:
Jresponse = Left(Jresponse, Pointer - 1)

'we take " out
Length = Len(Jresponse)
If Not IsNumeric(Jresponse) Then
Jresponse = Left(Jresponse, Length - 1)
Jresponse = Right(Jresponse, Length - 2)
End If

If MultiValue = True Then
    MultiValueJresponsePivot = Jresponse
    While InStr(MultiValueJresponse, SubFieldName) <> 0
        Length = Len(MultiValueJresponse)
        MultiValuePointer = InStr(MultiValueJresponse, SubFieldName)
        Jresponse = Right(MultiValueJresponse, Length - MultiValuePointer)
        MultiValuePointer = InStr(Jresponse, ":") + MultiValuePointer
        Jresponse = Right(Jresponse, Length - MultiValuePointer)
        MultiValueJresponse = Jresponse
        MultiValuePointer = InStr(Jresponse, ",")
        If InStr(Jresponse, ",") = 0 Then MultiValuePointer = InStr(Jresponse, "}")
        Jresponse = Left(Jresponse, MultiValuePointer - 1)
        Length = Len(Jresponse)
        Jresponse = Left(Jresponse, Length - 1)
        Jresponse = Right(Jresponse, Length - 2)
        MultiValueJresponsePivot = MultiValueJresponsePivot & ", " & Jresponse
    Wend
    Jresponse = MultiValueJresponsePivot
End If

If FieldisNull = True Then Jresponse = ""
FieldSeekerMultiValue = Jresponse


Exit Function

NoComa:
Pointer = InStr(Jresponse, "}")
GoTo InCaseofError

End Function

'ONLY WORKS FOR SUBTASKS, have to remove a few lines to make it work with main tickets
'Must search a single SUBTASK without search parameters,
'this one gets the current status of an issue,
'Will work only for status
Public Function CurrentStatusSeeker(Jresponse As String) As String

Dim Pointer As Long

Pointer = InStr(Jresponse, "{") 'Take out first expand bracket
Jresponse = Right(Jresponse, ResponseLength - Pointer)

Pointer = InStr(Jresponse, "{") + Pointer 'Take out issue bracket (this function only works for single issue searches)
Jresponse = Right(Jresponse, ResponseLength - Pointer)

Pointer = InStr(Jresponse, "{") + Pointer 'Take out fields bracket
Jresponse = Right(Jresponse, ResponseLength - Pointer)

Pointer = InStr(Jresponse, """parent""") + Pointer - 1 'remove to work with main task ********
Jresponse = Right(Jresponse, ResponseLength - Pointer)

Pointer = InStr(Jresponse, "}") + Pointer 'remove to work with main task *********
Jresponse = Right(Jresponse, ResponseLength - Pointer)

Pointer = InStr(Jresponse, """status""") + Pointer - 1 'We keep the field
Jresponse = Right(Jresponse, ResponseLength - Pointer)

Pointer = InStr(Jresponse, """name""") + Pointer 'Take out fields bracket
Jresponse = Right(Jresponse, ResponseLength - Pointer)

Pointer = InStr(Jresponse, ":") + Pointer 'we take out the field name
Jresponse = Right(Jresponse, ResponseLength - Pointer - 1)



On Error GoTo NoComa

Pointer = InStr(Jresponse, ",") 'We keep only the field value
InCaseofError:
Jresponse = Left(Jresponse, Pointer - 2)

CurrentStatusSeeker = Jresponse

Exit Function

NoComa:
Pointer = InStr(Jresponse, "}")
GoTo InCaseofError

End Function


Option Explicit
Dim JRA As JiraRestApi



' function called by the form when all the data is filled
Public Sub importIssuesToJira(jiraUser As String, Password As String, Projectkey As String, SysTicket As String)
    Set JRA = New JiraRestApi
    Dim ActiveSheetName As String

JRA.Username = jiraUser
JRA.Password = Password
JRA.IntranetLink = "http://eudca-jira01/jira/rest/api/2/"
JRA.Issuekey = Projectkey & "-" & SysTicket
JRA.Projectkey = Projectkey
JRA.ParentIssue = Projectkey & "-" & SysTicket
    
    ActiveSheetName = ActiveSheet.Name
    getIssuesFromSheet ActiveSheetName
End Sub



' function that makes a loop around all the intoduced issues and update each of it to jira
Private Sub getIssuesFromSheet(Sheet As String)
    Dim i As Long
    Dim Pointer As Long
    Dim summary As String
    Dim description As String
    Dim assignee As String
    Dim EffortInHours As String
    Dim Status As String
    Dim ResolutionDate As String
    Dim Issuetype As String
    Dim DmaicPhase As String
    Dim DueDate As String
    Dim LearWorkOrder As String
    Dim NM As String
    Dim LearHWDeliverables As String
    Dim LearHWTaskType As String
    Dim DueYear As String
    Dim DueMonth As String
    Dim DueDay As String
    Dim StartingDate As String
    Dim JiraPriority As String
    Dim Dayint As Integer
    Dim Monthint As Integer
    Dim CloseDatePivot As String
    Dim ActualEffortPivot As Variant
    Dim DeltaEffortPivot As Variant
    Dim LearSWDeliverables As String
    Dim LearSWDTaskType As String
    Dim LearCOREDeliverables As String
    Dim LearCORETaskType As String
    Dim LearMECHDeliverables As String
    Dim LearMECHTaskType As String
    Dim LearPCBDeliverables As String
    Dim LearPCBTaskType As String
    Dim LearSYSDeliverables As String
    Dim LearSYSTaskType As String
    Dim LearSAFETYMNGDeliverables As String
    Dim LearSAFETYMNGTaskType As String
    
'rest api fields
    Dim ProjectField As String
    Dim ParentField As String
    Dim SummaryField As String
    Dim StartingDateField As String
    Dim IssuetypeField As String
    Dim DuedateField As String
    Dim LearWorkOrderField As String
    Dim NMField As String
    Dim LearHWDeliverablesField As String
    Dim LearHWTaskTypeField As String
    Dim DescriptionField As String
    Dim OriginalestimateField As String
    Dim DmaicPhaseField As String
    Dim AssigneeField As String
    Dim PriorityField As String
    Dim ProjectCompletionField As String
    Dim DataFields As String
    
    Dim DatePivot As Date
    Dim CloseDate As Date
    Dim NowDate As Date
    
    ' variables for indentifying the columns
    Dim summaryCol As Long
    Dim SystemmicNumberCol As Long
    Dim issueStatusCol As Long
    Dim SendToJira_Col As Long
    Dim JiraAssignee_Col As Long
    Dim EffortInHours_Col As Long
    Dim DmaicPhase_Col As Long
    Dim CIPTaskDescription_Col As Long
    Dim LastRowCheck As Long
    Dim YearPivot_Col As Long
    Dim MonthPivot_Col As Long
    Dim DayPivot_Col As Long
    Dim CompletionDateJira_Col As Long
    Dim StartingDate_Col As Long
    Dim JiraPriority_Col As Long
    Dim FinalROW_Row As Long
    Dim TaskCompletion_Col As Long
    Dim DueDate_Col As Long
    Dim DueDateCW_Col As Long
    Dim ActualEffort_Col As Long
    Dim DeltaEffort_Col As Long
    Dim TaskDone_Col As Long
    Dim SWDeliverables_Col As Long
    Dim SWDTaskType_Col As Long
    Dim COREDeliverables_Col As Long
    Dim CORETaskType_Col As Long
    Dim MECHDeliverables_Col As Long
    Dim MECHTaskType_Col As Long
    Dim PCBDeliverables_Col As Long
    Dim PCBTaskType_Col As Long
    Dim SYSDeliverables_Col As Long
    Dim SYSTaskType_Col As Long
    Dim SAFETYMNGDeliverables_Col As Long
    Dim SAFETYMNGTaskType_Col As Long
    
    'variable temporal
    Dim tempKey As String
    
    'Pivot vars for JiraSenddata
    
    
    'Defining the column number for the info we need to check
    
    summaryCol = Sheets(Sheet).Names("JiraSummaryCol").RefersToRange.Column
    SendToJira_Col = Sheets(Sheet).Names("SendToJira").RefersToRange.Column
    JiraAssignee_Col = Sheets(Sheet).Names("JiraAssigneeCol").RefersToRange.Column
    EffortInHours_Col = Sheets(Sheet).Names("EstimatedEffort").RefersToRange.Column
    DmaicPhase_Col = Sheets(Sheet).Names("DmaicPhase").RefersToRange.Column
    CIPTaskDescription_Col = Sheets(Sheet).Names("CIPTaskDescription").RefersToRange.Column
    LastRowCheck = Sheets(Sheet).Names("Task_NameCol").RefersToRange.Column
    SWDeliverables_Col = ActiveSheet.Names("SW_Dleliverables").RefersToRange.Column
    SWDTaskType_Col = ActiveSheet.Names("SW_TaskType").RefersToRange.Column
    DueDate_Col = ActiveSheet.Names("DueDate").RefersToRange.Column
    DueDateCW_Col = ActiveSheet.Names("DueDateCW").RefersToRange.Column
    StartingDate_Col = ActiveSheet.Names("StartingDate").RefersToRange.Column
    JiraPriority_Col = ActiveSheet.Names("JiraPriority").RefersToRange.Column
    FinalROW_Row = ActiveSheet.Names("Final_Row").RefersToRange.Row
    COREDeliverables_Col = ActiveSheet.Names("Core_Deliverables").RefersToRange.Column
    CORETaskType_Col = ActiveSheet.Names("COre_TaskType").RefersToRange.Column
    MECHDeliverables_Col = ActiveSheet.Names("LEAR_MECH_Deliverables").RefersToRange.Column
    MECHTaskType_Col = ActiveSheet.Names("LEAR_MECH_Task_Type").RefersToRange.Column
    PCBDeliverables_Col = ActiveSheet.Names("LEAR_PCB_Deliverables").RefersToRange.Column
    PCBTaskType_Col = ActiveSheet.Names("LEAR_PCB_Task_Type").RefersToRange.Column
    SYSDeliverables_Col = ActiveSheet.Names("LEAR_SYS_Deliverables").RefersToRange.Column
    SYSTaskType_Col = ActiveSheet.Names("LEAR_SYS_Task_Type").RefersToRange.Column
    SAFETYMNGDeliverables_Col = ActiveSheet.Names("LEAR_SAFETY_MNG_Deliverables").RefersToRange.Column
    SAFETYMNGTaskType_Col = ActiveSheet.Names("LEAR_SAFETY_MNG_Task_Type").RefersToRange.Column
    
    'Defining the first row to loop on the planning grid
   i = Sheets(Sheet).Names("Task_NameCol").RefersToRange.Row + 1
   


   
   'Stop looping once you reach the end of the planning grid
    While (Sheets(Sheet).Cells(i, LastRowCheck).Value <> "Final ROW")
        ' if the line contains a yes, send to jira
        While (Sheets(Sheet).Cells(i, SendToJira_Col).Value = "Yes")

             
                 summary = Sheets(Sheet).Cells(i, summaryCol).Value
                 assignee = Sheets(Sheet).Cells(i, JiraAssignee_Col).Value
                 description = Sheets(Sheet).Cells(i, CIPTaskDescription_Col).Value
                 DueDate = Sheets(Sheet).Cells(i, DueDate_Col).text
                 StartingDate = Sheets(Sheet).Cells(i, StartingDate_Col).text
                 LearWorkOrder = Sheets(Sheet).Cells(i, JiraPriority_Col).Value
                 NM = Sheets(Sheet).Cells(i, DmaicPhase_Col).Value
                 LearHWDeliverables = Sheets(Sheet).Cells(i, DueDateCW_Col).Value
                 LearHWTaskType = Sheets(Sheet).Cells(i, EffortInHours_Col).Value
                 JiraPriority = 3
                 LearSWDeliverables = Sheets(Sheet).Cells(i, SWDeliverables_Col).Value
                 LearSWDTaskType = Sheets(Sheet).Cells(i, SWDTaskType_Col).Value
                 LearCOREDeliverables = Sheets(Sheet).Cells(i, COREDeliverables_Col).Value
                 LearCORETaskType = Sheets(Sheet).Cells(i, CORETaskType_Col).Value
                 LearMECHDeliverables = Sheets(Sheet).Cells(i, MECHDeliverables_Col).Value
                 LearMECHTaskType = Sheets(Sheet).Cells(i, MECHTaskType_Col).Value
                 LearPCBDeliverables = Sheets(Sheet).Cells(i, PCBDeliverables_Col).Value
                 LearPCBTaskType = Sheets(Sheet).Cells(i, PCBTaskType_Col).Value
                 LearSYSDeliverables = Sheets(Sheet).Cells(i, SYSDeliverables_Col).Value
                 LearSYSTaskType = Sheets(Sheet).Cells(i, SYSTaskType_Col).Value
                 LearSAFETYMNGDeliverables = Sheets(Sheet).Cells(i, SAFETYMNGDeliverables_Col).Value
                 LearSAFETYMNGTaskType = Sheets(Sheet).Cells(i, SAFETYMNGTaskType_Col).Value
                 
                Dim Name As String
                Range("E4").Select
                ActiveCell.Select
                Name = ActiveCell.Value

                Select Case Name

                Case "HW"
                Issuetype = 10102
                 'We launch the function that creates the issue on jira
                'Create Jira Issue
                
                'Link to see how to implement data in different kind of fields---->
                'https://developer.atlassian.com/display/JIRADEV/JIRA+REST+API+Example+-+Create+Issue
                ProjectField = " ""project"" : { ""key"" : """ & JRA.Projectkey & """ }, " 'project key name
                AssigneeField = " ""assignee"" : { ""name"" : """ & assignee & """ }, " 'assignee
                PriorityField = " ""priority"" : {""id"" : """ & JiraPriority & """}, " 'priority
                SummaryField = " ""summary"" : """ & summary & """, " 'Summary field
                StartingDateField = " ""customfield_13031"" : """ & StartingDate & """, " 'Starting Date
                LearWorkOrderField = " ""customfield_12633"" : """ & summary & """, " ' Lear Work Order
                NMField = " ""customfield_15530"" : """ & NM & """, " 'Summary field
                LearHWDeliverablesField = " ""customfield_13535"" : [{ ""value"": """ & LearHWDeliverables & """ }], "
                LearHWTaskTypeField = " ""customfield_13534"" : { ""value"": """ & LearHWTaskType & """ }, " 'NM Value
                IssuetypeField = " ""issuetype"" : { ""id"" : """ & Issuetype & """ }, " 'Issue type id
                DuedateField = " ""duedate"" : """ & DueDate & """ " 'Due date in "YYYY-MM-DD" Format
                
                DataFields = ProjectField & AssigneeField & PriorityField & SummaryField & StartingDateField & LearWorkOrderField & _
                NMField & LearHWDeliverablesField & LearHWTaskTypeField & IssuetypeField & DuedateField 'Merge fields together
                
                JRA.SendData = " { ""fields"" : {  " & DataFields & "  } } " 'Merge fields with the header
                
                Call JRA.CreateTicket
                
                
                ActiveSheet.Cells(65, 5).Value = JRA.Response
                i = i + 1
                
                Case "SW"
                Issuetype = 10200
                 'We launch the function that creates the issue on jira
                'Create Jira Issue
                
                'Link to see how to implement data in different kind of fields---->
                'https://developer.atlassian.com/display/JIRADEV/JIRA+REST+API+Example+-+Create+Issue
                ProjectField = " ""project"" : { ""key"" : """ & JRA.Projectkey & """ }, " 'project key name
                AssigneeField = " ""assignee"" : { ""name"" : """ & assignee & """ }, " 'assignee
                PriorityField = " ""priority"" : {""id"" : """ & JiraPriority & """}, " 'priority
                SummaryField = " ""summary"" : """ & summary & """, " 'Summary field
                StartingDateField = " ""customfield_13031"" : """ & StartingDate & """, " 'Starting Date
                LearWorkOrderField = " ""customfield_12633"" : """ & summary & """, " ' Lear Work Order
                NMField = " ""customfield_15530"" : """ & NM & """, " 'Summary field
                LearHWDeliverablesField = " ""customfield_13630"" : [{ ""value"": """ & LearSWDeliverables & """ }], "
                LearHWTaskTypeField = " ""customfield_13631"" : { ""value"": """ & LearSWDTaskType & """ }, " 'NM Value
                IssuetypeField = " ""issuetype"" : { ""id"" : """ & Issuetype & """ }, " 'Issue type id
                DuedateField = " ""duedate"" : """ & DueDate & """ " 'Due date in "YYYY-MM-DD" Format
                
                DataFields = ProjectField & AssigneeField & PriorityField & SummaryField & StartingDateField & LearWorkOrderField & _
                NMField & LearHWDeliverablesField & LearHWTaskTypeField & IssuetypeField & DuedateField 'Merge fields together
                
                JRA.SendData = " { ""fields"" : {  " & DataFields & "  } } " 'Merge fields with the header
                
                Call JRA.CreateTicket
                
                
                ActiveSheet.Cells(65, 5).Value = JRA.Response
                i = i + 1
                
                
                Case "CORE"
                Issuetype = 10101
                 'We launch the function that creates the issue on jira
                'Create Jira Issue
                
                'Link to see how to implement data in different kind of fields---->
                'https://developer.atlassian.com/display/JIRADEV/JIRA+REST+API+Example+-+Create+Issue
                ProjectField = " ""project"" : { ""key"" : """ & JRA.Projectkey & """ }, " 'project key name
                AssigneeField = " ""assignee"" : { ""name"" : """ & assignee & """ }, " 'assignee
                PriorityField = " ""priority"" : {""id"" : """ & JiraPriority & """}, " 'priority
                SummaryField = " ""summary"" : """ & summary & """, " 'Summary field
                StartingDateField = " ""customfield_13031"" : """ & StartingDate & """, " 'Starting Date
                LearWorkOrderField = " ""customfield_12633"" : """ & summary & """, " ' Lear Work Order
                NMField = " ""customfield_15530"" : """ & NM & """, " 'Summary field
                LearHWDeliverablesField = " ""customfield_13430"" : [{ ""value"": """ & LearCOREDeliverables & """ }], "
                LearHWTaskTypeField = " ""customfield_13531"" : { ""value"": """ & LearCORETaskType & """ }, " 'NM Value
                IssuetypeField = " ""issuetype"" : { ""id"" : """ & Issuetype & """ }, " 'Issue type id
                DuedateField = " ""duedate"" : """ & DueDate & """ " 'Due date in "YYYY-MM-DD" Format
                
                DataFields = ProjectField & AssigneeField & PriorityField & SummaryField & StartingDateField & LearWorkOrderField & _
                NMField & LearHWDeliverablesField & LearHWTaskTypeField & IssuetypeField & DuedateField 'Merge fields together
                
                JRA.SendData = " { ""fields"" : {  " & DataFields & "  } } " 'Merge fields with the header
                
                Call JRA.CreateTicket
                
                
                ActiveSheet.Cells(65, 5).Value = JRA.Response
                i = i + 1
                
                
                
                Case "MECH"
                Issuetype = 10100
                 'We launch the function that creates the issue on jira
                'Create Jira Issue
                
                'Link to see how to implement data in different kind of fields---->
                'https://developer.atlassian.com/display/JIRADEV/JIRA+REST+API+Example+-+Create+Issue
                ProjectField = " ""project"" : { ""key"" : """ & JRA.Projectkey & """ }, " 'project key name
                AssigneeField = " ""assignee"" : { ""name"" : """ & assignee & """ }, " 'assignee
                PriorityField = " ""priority"" : {""id"" : """ & JiraPriority & """}, " 'priority
                SummaryField = " ""summary"" : """ & summary & """, " 'Summary field
                StartingDateField = " ""customfield_13031"" : """ & StartingDate & """, " 'Starting Date
                LearWorkOrderField = " ""customfield_12633"" : """ & summary & """, " ' Lear Work Order
                NMField = " ""customfield_15530"" : """ & NM & """, " 'Summary field
                LearHWDeliverablesField = " ""customfield_13533"" : [{ ""value"": """ & LearMECHDeliverables & """ }], "
                LearHWTaskTypeField = " ""customfield_13532"" : { ""value"": """ & LearMECHTaskType & """ }, " 'NM Value
                IssuetypeField = " ""issuetype"" : { ""id"" : """ & Issuetype & """ }, " 'Issue type id
                DuedateField = " ""duedate"" : """ & DueDate & """ " 'Due date in "YYYY-MM-DD" Format
                
                DataFields = ProjectField & AssigneeField & PriorityField & SummaryField & StartingDateField & LearWorkOrderField & _
                NMField & LearHWDeliverablesField & LearHWTaskTypeField & IssuetypeField & DuedateField 'Merge fields together
                
                JRA.SendData = " { ""fields"" : {  " & DataFields & "  } } " 'Merge fields with the header
                
                Call JRA.CreateTicket
                
                
                ActiveSheet.Cells(65, 5).Value = JRA.Response
                i = i + 1
                
                Case "PCB"
                Issuetype = 10001
                 'We launch the function that creates the issue on jira
                'Create Jira Issue
                
                'Link to see how to implement data in different kind of fields---->
                'https://developer.atlassian.com/display/JIRADEV/JIRA+REST+API+Example+-+Create+Issue
                ProjectField = " ""project"" : { ""key"" : """ & JRA.Projectkey & """ }, " 'project key name
                AssigneeField = " ""assignee"" : { ""name"" : """ & assignee & """ }, " 'assignee
                PriorityField = " ""priority"" : {""id"" : """ & JiraPriority & """}, " 'priority
                SummaryField = " ""summary"" : """ & summary & """, " 'Summary field
                StartingDateField = " ""customfield_13031"" : """ & StartingDate & """, " 'Starting Date
                LearWorkOrderField = " ""customfield_12633"" : """ & summary & """, " ' Lear Work Order
                NMField = " ""customfield_15530"" : """ & NM & """, " 'Summary field
                LearHWDeliverablesField = " ""customfield_13033"" : [{ ""value"": """ & LearPCBDeliverables & """ }], "
                LearHWTaskTypeField = " ""customfield_13530"" : { ""value"": """ & LearPCBTaskType & """ }, " 'NM Value
                IssuetypeField = " ""issuetype"" : { ""id"" : """ & Issuetype & """ }, " 'Issue type id
                DuedateField = " ""duedate"" : """ & DueDate & """ " 'Due date in "YYYY-MM-DD" Format
                
                DataFields = ProjectField & AssigneeField & PriorityField & SummaryField & StartingDateField & LearWorkOrderField & _
                NMField & LearHWDeliverablesField & LearHWTaskTypeField & IssuetypeField & DuedateField 'Merge fields together
                
                JRA.SendData = " { ""fields"" : {  " & DataFields & "  } } " 'Merge fields with the header
                
                Call JRA.CreateTicket
                
                
                ActiveSheet.Cells(65, 5).Value = JRA.Response
                i = i + 1
                
                
                Case "SYS"
                Issuetype = 11500
                 'We launch the function that creates the issue on jira
                'Create Jira Issue
                
                'Link to see how to implement data in different kind of fields---->
                'https://developer.atlassian.com/display/JIRADEV/JIRA+REST+API+Example+-+Create+Issue
                ProjectField = " ""project"" : { ""key"" : """ & JRA.Projectkey & """ }, " 'project key name
                AssigneeField = " ""assignee"" : { ""name"" : """ & assignee & """ }, " 'assignee
                PriorityField = " ""priority"" : {""id"" : """ & JiraPriority & """}, " 'priority
                SummaryField = " ""summary"" : """ & summary & """, " 'Summary field
                StartingDateField = " ""customfield_13031"" : """ & StartingDate & """, " 'Starting Date
                LearWorkOrderField = " ""customfield_12633"" : """ & summary & """, " ' Lear Work Order
                NMField = " ""customfield_15530"" : """ & NM & """, " 'Summary field
                LearHWDeliverablesField = " ""customfield_16731"" : [{ ""value"": """ & LearSYSDeliverables & """ }], "
                LearHWTaskTypeField = " ""customfield_16730"" : { ""value"": """ & LearSYSTaskType & """ }, " 'NM Value
                IssuetypeField = " ""issuetype"" : { ""id"" : """ & Issuetype & """ }, " 'Issue type id
                DuedateField = " ""duedate"" : """ & DueDate & """ " 'Due date in "YYYY-MM-DD" Format
                
                DataFields = ProjectField & AssigneeField & PriorityField & SummaryField & StartingDateField & LearWorkOrderField & _
                NMField & LearHWDeliverablesField & LearHWTaskTypeField & IssuetypeField & DuedateField 'Merge fields together
                
                JRA.SendData = " { ""fields"" : {  " & DataFields & "  } } " 'Merge fields with the header
                
                Call JRA.CreateTicket
                
                
                ActiveSheet.Cells(65, 5).Value = JRA.Response
                i = i + 1
                
                Case "SAFETY MNG"
                Issuetype = 11400
                 'We launch the function that creates the issue on jira
                'Create Jira Issue
                
                'Link to see how to implement data in different kind of fields---->
                'https://developer.atlassian.com/display/JIRADEV/JIRA+REST+API+Example+-+Create+Issue
                ProjectField = " ""project"" : { ""key"" : """ & JRA.Projectkey & """ }, " 'project key name
                AssigneeField = " ""assignee"" : { ""name"" : """ & assignee & """ }, " 'assignee
                PriorityField = " ""priority"" : {""id"" : """ & JiraPriority & """}, " 'priority
                SummaryField = " ""summary"" : """ & summary & """, " 'Summary field
                StartingDateField = " ""customfield_13031"" : """ & StartingDate & """, " 'Starting Date
                LearWorkOrderField = " ""customfield_12633"" : """ & summary & """, " ' Lear Work Order
                NMField = " ""customfield_15530"" : """ & NM & """, " 'Summary field
                LearHWDeliverablesField = " ""customfield_16632"" : [{ ""value"": """ & LearSAFETYMNGDeliverables & """ }], "
                LearHWTaskTypeField = " ""customfield_16630"" : { ""value"": """ & LearSAFETYMNGTaskType & """ }, " 'NM Value
                IssuetypeField = " ""issuetype"" : { ""id"" : """ & Issuetype & """ }, " 'Issue type id
                DuedateField = " ""duedate"" : """ & DueDate & """ " 'Due date in "YYYY-MM-DD" Format
                
                DataFields = ProjectField & AssigneeField & PriorityField & SummaryField & StartingDateField & LearWorkOrderField & _
                NMField & LearHWDeliverablesField & LearHWTaskTypeField & IssuetypeField & DuedateField 'Merge fields together
                
                JRA.SendData = " { ""fields"" : {  " & DataFields & "  } } " 'Merge fields with the header
                
                Call JRA.CreateTicket
                
                
                ActiveSheet.Cells(65, 5).Value = JRA.Response
                i = i + 1
                 
                End Select
         
                
         Wend
         i = i + 1
    Wend
   
End Sub


Option Explicit
' Developed by Contextures Inc.
' www.contextures.com
Private Sub Worksheet_Change(ByVal Target As Range)
Dim rngDV As Range
Dim oldVal As String
Dim newVal As String
If Target.Count > 1 Then GoTo exitHandler

On Error Resume Next
Set rngDV = Cells.SpecialCells(xlCellTypeAllValidation)
On Error GoTo exitHandler

If rngDV Is Nothing Then GoTo exitHandler

If Intersect(Target, rngDV) Is Nothing Then
   'do nothing
Else
  Application.EnableEvents = False
  newVal = Target.Value
  Application.Undo
  oldVal = Target.Value
  Target.Value = newVal
  
  If Target.Column = 25 Then
    If oldVal = "" Then
      'do nothing
    Else
      If newVal = "" Then
        'do nothing
      Else
        Target.Value = oldVal _
          & ", " & newVal
      End If
    End If
  End If
  
  If Target.Column = 23 Then
    If oldVal = "" Then
      'do nothing
    Else
      If newVal = "" Then
        'do nothing
      Else
        Target.Value = oldVal _
          & ", " & newVal
      End If
    End If
  End If

If Target.Column = 21 Then
    If oldVal = "" Then
      'do nothing
    Else
      If newVal = "" Then
        'do nothing
      Else
        Target.Value = oldVal _
          & ", " & newVal
      End If
    End If
  End If
  
  If Target.Column = 19 Then
    If oldVal = "" Then
      'do nothing
    Else
      If newVal = "" Then
        'do nothing
      Else
        Target.Value = oldVal _
          & ", " & newVal
      End If
    End If
  End If
  
  If Target.Column = 17 Then
    If oldVal = "" Then
      'do nothing
    Else
      If newVal = "" Then
        'do nothing
      Else
        Target.Value = oldVal _
          & ", " & newVal
      End If
    End If
  End If

If Target.Column = 15 Then
    If oldVal = "" Then
      'do nothing
    Else
      If newVal = "" Then
        'do nothing
      Else
        Target.Value = oldVal _
          & ", " & newVal
      End If
    End If
  End If
End If

exitHandler:
  Application.EnableEvents = True
End Sub


