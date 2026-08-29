Attribute VB_Name = "Generate_HRAnalysis"
Option Explicit

'========================================================
' HR ANALYTICS PORTFOLIO
' HR ANALYSIS
'========================================================

Public Sub GenerateHRAnalysis()

On Error GoTo ErrorHandler

Dim wsData As Worksheet
Dim wsAnalysis As Worksheet

Dim lastRow As Long
Dim r As Long

Dim totalEmployees As Long
Dim activeEmployees As Long
Dim inactiveEmployees As Long
Dim onLeaveEmployees As Long

Dim totalSalary As Double
Dim averageSalary As Double
Dim highestSalary As Double
Dim lowestSalary As Double

Dim performanceTotal As Double
Dim performanceCount As Long
Dim averagePerformance As Double

Dim admissionDate As Date
Dim tenureYears As Double
Dim totalTenure As Double
Dim tenureCount As Long
Dim longestTenure As Double

Dim statusValue As String
Dim performanceValue As Variant
Dim salaryValue As Double

Dim rowTenure As Long
Dim rowDepartment As Long
Dim rowRisk As Long
Dim lastDepartmentRow As Long
Dim lastRiskRow As Long

Call Optimization

'====================================================
' SOURCE SHEET
'====================================================

On Error Resume Next
Set wsData = ThisWorkbook.Worksheets("EmployeeData")
On Error GoTo ErrorHandler

If wsData Is Nothing Then

    MsgBox "The 'EmployeeData' sheet was not found." & vbCrLf & vbCrLf & _
           "Please import the employee base first.", _
           vbExclamation, "HR Analysis"

    Call Desoptimization
    Exit Sub

End If

'====================================================
' CREATE OR CLEAR ANALYSIS SHEET
'====================================================

On Error Resume Next
Set wsAnalysis = ThisWorkbook.Worksheets("HR Analysis")
On Error GoTo ErrorHandler

If wsAnalysis Is Nothing Then

    Set wsAnalysis = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))

    wsAnalysis.Name = "HR Analysis"

Else

    wsAnalysis.Cells.Clear

End If

'====================================================
' FIND LAST ROW
'====================================================

lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).row

If lastRow < 2 Then

    MsgBox "No employee records were found in EmployeeData.", _
           vbExclamation, "HR Analysis"

    Call Desoptimization
    Exit Sub

End If

'====================================================
' INITIAL VALUES
'====================================================

highestSalary = 0
lowestSalary = 0

'====================================================
' READ EMPLOYEE DATA
'====================================================

For r = 2 To lastRow

    totalEmployees = totalEmployees + 1

    statusValue = LCase(Trim(CStr(wsData.Cells(r, 7).Value)))

    Select Case statusValue

        Case "ativo"
            activeEmployees = activeEmployees + 1

        Case "afastado"
            onLeaveEmployees = onLeaveEmployees + 1

        Case "desligado"
            inactiveEmployees = inactiveEmployees + 1

    End Select

    '--------------------------------------------
    ' Salary
    '--------------------------------------------

    If IsNumeric(wsData.Cells(r, 5).Value) Then

        salaryValue = CDbl(wsData.Cells(r, 5).Value)

        totalSalary = totalSalary + salaryValue

        If highestSalary = 0 Or salaryValue > highestSalary Then
            highestSalary = salaryValue
        End If

        If lowestSalary = 0 Or salaryValue < lowestSalary Then
            lowestSalary = salaryValue
        End If

    End If

    '--------------------------------------------
    ' Performance
    '--------------------------------------------

    performanceValue = wsData.Cells(r, 8).Value

    If IsNumeric(performanceValue) Then

        performanceTotal = performanceTotal + CDbl(performanceValue)
        performanceCount = performanceCount + 1

    End If

    '--------------------------------------------
    ' Tenure
    '--------------------------------------------

    If IsDate(wsData.Cells(r, 6).Value) Then

        admissionDate = CDate(wsData.Cells(r, 6).Value)

        tenureYears = DateDiff("d", admissionDate, Date) / 365.25

        totalTenure = totalTenure + tenureYears
        tenureCount = tenureCount + 1

        If tenureYears > longestTenure Then
            longestTenure = tenureYears
        End If

    End If

Next r

'====================================================
' CALCULATE AVERAGES
'====================================================

If totalEmployees > 0 Then
    averageSalary = totalSalary / totalEmployees
End If

If performanceCount > 0 Then
    averagePerformance = performanceTotal / performanceCount
End If

'====================================================
' TITLE
'====================================================

With wsAnalysis

    .Range("A1:J1").Merge
    .Range("A1").Value = "HR ANALYTICS"

    .Range("A2:J2").Merge
    .Range("A2").Value = "Workforce Overview and Employee Analysis"

End With

'====================================================
' WORKFORCE OVERVIEW
'====================================================

wsAnalysis.Range("A4:B4").Merge
wsAnalysis.Range("A4").Value = "WORKFORCE OVERVIEW"

wsAnalysis.Range("A5").Value = "Metric"
wsAnalysis.Range("B5").Value = "Value"

wsAnalysis.Range("A6").Value = "Total Employees"
wsAnalysis.Range("B6").Value = totalEmployees

wsAnalysis.Range("A7").Value = "Active Employees"
wsAnalysis.Range("B7").Value = activeEmployees

wsAnalysis.Range("A8").Value = "Employees on Leave"
wsAnalysis.Range("B8").Value = onLeaveEmployees

wsAnalysis.Range("A9").Value = "Former Employees"
wsAnalysis.Range("B9").Value = inactiveEmployees

'====================================================
' SALARY ANALYSIS
'====================================================

wsAnalysis.Range("D4:E4").Merge
wsAnalysis.Range("D4").Value = "SALARY ANALYSIS"

wsAnalysis.Range("D5").Value = "Metric"
wsAnalysis.Range("E5").Value = "Value"

wsAnalysis.Range("D6").Value = "Total Payroll"
wsAnalysis.Range("E6").Value = totalSalary

wsAnalysis.Range("D7").Value = "Average Salary"
wsAnalysis.Range("E7").Value = averageSalary

wsAnalysis.Range("D8").Value = "Highest Salary"
wsAnalysis.Range("E8").Value = highestSalary

wsAnalysis.Range("D9").Value = "Lowest Salary"
wsAnalysis.Range("E9").Value = lowestSalary

wsAnalysis.Range("D10").Value = "Salary Range"
wsAnalysis.Range("E10").Value = highestSalary - lowestSalary

'====================================================
' PERFORMANCE ANALYSIS
'====================================================

wsAnalysis.Range("G4:H4").Merge
wsAnalysis.Range("G4").Value = "PERFORMANCE ANALYSIS"

wsAnalysis.Range("G5").Value = "Metric"
wsAnalysis.Range("H5").Value = "Value"

wsAnalysis.Range("G6").Value = "Average Performance"
wsAnalysis.Range("H6").Value = averagePerformance

wsAnalysis.Range("G7").Value = "Employees Rated 5"
wsAnalysis.Range("H7").Value = CountPerformance(wsData, 5, lastRow)

wsAnalysis.Range("G8").Value = "Employees Rated 4"
wsAnalysis.Range("H8").Value = CountPerformance(wsData, 4, lastRow)

wsAnalysis.Range("G9").Value = "Employees Rated 3"
wsAnalysis.Range("H9").Value = CountPerformance(wsData, 3, lastRow)

wsAnalysis.Range("G10").Value = "Employees Rated 1-2"
wsAnalysis.Range("H10").Value = CountLowPerformance(wsData, lastRow)

'====================================================
' TENURE ANALYSIS
' UNIT: YEARS
'====================================================

rowTenure = 13

wsAnalysis.Range("A" & rowTenure & ":B" & rowTenure).Merge
wsAnalysis.Cells(rowTenure, 1).Value = "TENURE ANALYSIS (YEARS)"

wsAnalysis.Cells(rowTenure + 1, 1).Value = "Metric"
wsAnalysis.Cells(rowTenure + 1, 2).Value = "Value"

wsAnalysis.Cells(rowTenure + 2, 1).Value = "Average Tenure"

If tenureCount > 0 Then
    wsAnalysis.Cells(rowTenure + 2, 2).Value = totalTenure / tenureCount
Else
    wsAnalysis.Cells(rowTenure + 2, 2).Value = 0
End If

wsAnalysis.Cells(rowTenure + 3, 1).Value = "Longest Tenure"
wsAnalysis.Cells(rowTenure + 3, 2).Value = longestTenure

wsAnalysis.Cells(rowTenure + 4, 1).Value = "Less Than 1 Year"
wsAnalysis.Cells(rowTenure + 4, 2).Value = CountTenureLessThan(wsData, lastRow, 1)

wsAnalysis.Cells(rowTenure + 5, 1).Value = "1-3 Years"
wsAnalysis.Cells(rowTenure + 5, 2).Value = CountTenureBetween(wsData, lastRow, 1, 3)

wsAnalysis.Cells(rowTenure + 6, 1).Value = "3-5 Years"
wsAnalysis.Cells(rowTenure + 6, 2).Value = CountTenureBetween(wsData, lastRow, 3, 5)

wsAnalysis.Cells(rowTenure + 7, 1).Value = "More Than 5 Years"
wsAnalysis.Cells(rowTenure + 7, 2).Value = CountTenureMoreThan(wsData, lastRow, 5)

'====================================================
' DEPARTMENT ANALYSIS
'====================================================

rowDepartment = 13

wsAnalysis.Range( _
    "D" & rowDepartment & ":G" & rowDepartment).Merge

wsAnalysis.Cells(rowDepartment, 4).Value = "DEPARTMENT ANALYSIS"

wsAnalysis.Cells(rowDepartment + 1, 4).Value = "Department"
wsAnalysis.Cells(rowDepartment + 1, 5).Value = "Employees"
wsAnalysis.Cells(rowDepartment + 1, 6).Value = "Avg Salary"
wsAnalysis.Cells(rowDepartment + 1, 7).Value = "Avg Performance"

Call BuildDepartmentAnalysis( _
    wsData, _
    wsAnalysis, _
    lastRow, _
    rowDepartment + 2)

lastDepartmentRow = wsAnalysis.Cells( _
    wsAnalysis.Rows.Count, 4).End(xlUp).row

'====================================================
' RISK ANALYSIS
'====================================================

rowRisk = lastDepartmentRow + 3

wsAnalysis.Range( _
    "A" & rowRisk & ":H" & rowRisk).Merge

wsAnalysis.Cells(rowRisk, 1).Value = "EMPLOYEE RISK ANALYSIS"

wsAnalysis.Cells(rowRisk + 1, 1).Value = "Employee ID"
wsAnalysis.Cells(rowRisk + 1, 2).Value = "Employee Name"
wsAnalysis.Cells(rowRisk + 1, 3).Value = "Department"
wsAnalysis.Cells(rowRisk + 1, 4).Value = "Salary"
wsAnalysis.Cells(rowRisk + 1, 5).Value = "Performance"
wsAnalysis.Cells(rowRisk + 1, 6).Value = "Tenure"
wsAnalysis.Cells(rowRisk + 1, 7).Value = "Risk Category"
wsAnalysis.Cells(rowRisk + 1, 8).Value = "Recommendation"

Call BuildRiskAnalysis( _
    wsData, _
    wsAnalysis, _
    lastRow, _
    rowRisk + 2)

lastRiskRow = wsAnalysis.Cells( _
    wsAnalysis.Rows.Count, 1).End(xlUp).row

'====================================================
' FORMATTING
'====================================================

Call FormatHRAnalysis( _
    wsAnalysis, _
    rowTenure, _
    rowDepartment, _
    rowRisk, _
    lastDepartmentRow, _
    lastRiskRow)

'====================================================
' FINISH
'====================================================

wsAnalysis.Activate

Call Desoptimization

MsgBox "HR Analysis successfully generated.", _
       vbInformation, "HR Analysis"

Exit Sub


ErrorHandler:


Call Desoptimization

MsgBox "An error occurred while generating HR Analysis:" & _
       vbCrLf & vbCrLf & Err.description, _
       vbCritical, "HR Analysis"


End Sub

'========================================================
' COUNT PERFORMANCE
'========================================================

Private Function CountPerformance( _
ByVal ws As Worksheet, _
ByVal rating As Double, _
ByVal lastRow As Long) As Long


Dim r As Long

For r = 2 To lastRow

    If IsNumeric(ws.Cells(r, 8).Value) Then

        If CDbl(ws.Cells(r, 8).Value) = rating Then
            CountPerformance = CountPerformance + 1
        End If

    End If

Next r


End Function

'========================================================
' COUNT LOW PERFORMANCE
'========================================================

Private Function CountLowPerformance( _
ByVal ws As Worksheet, _
ByVal lastRow As Long) As Long


Dim r As Long

For r = 2 To lastRow

    If IsNumeric(ws.Cells(r, 8).Value) Then

        If CDbl(ws.Cells(r, 8).Value) <= 2 Then
            CountLowPerformance = CountLowPerformance + 1
        End If

    End If

Next r


End Function

'========================================================
' COUNT TENURE LESS THAN
'========================================================

Private Function CountTenureLessThan( _
ByVal ws As Worksheet, _
ByVal lastRow As Long, _
ByVal years As Double) As Long


Dim r As Long
Dim tenure As Double

For r = 2 To lastRow

    If IsDate(ws.Cells(r, 6).Value) Then

        tenure = DateDiff( _
            "d", _
            CDate(ws.Cells(r, 6).Value), _
            Date) / 365.25

        If tenure < years Then
            CountTenureLessThan = CountTenureLessThan + 1
        End If

    End If

Next r


End Function

'========================================================
' COUNT TENURE BETWEEN
'========================================================

Private Function CountTenureBetween( _
ByVal ws As Worksheet, _
ByVal lastRow As Long, _
ByVal minYears As Double, _
ByVal maxYears As Double) As Long


Dim r As Long
Dim tenure As Double

For r = 2 To lastRow

    If IsDate(ws.Cells(r, 6).Value) Then

        tenure = DateDiff( _
            "d", _
            CDate(ws.Cells(r, 6).Value), _
            Date) / 365.25

        If tenure >= minYears And tenure < maxYears Then
            CountTenureBetween = CountTenureBetween + 1
        End If

    End If

Next r


End Function

'========================================================
' COUNT TENURE MORE THAN
'========================================================

Private Function CountTenureMoreThan( _
ByVal ws As Worksheet, _
ByVal lastRow As Long, _
ByVal years As Double) As Long


Dim r As Long
Dim tenure As Double

For r = 2 To lastRow

    If IsDate(ws.Cells(r, 6).Value) Then

        tenure = DateDiff( _
            "d", _
            CDate(ws.Cells(r, 6).Value), _
            Date) / 365.25

        If tenure >= years Then
            CountTenureMoreThan = CountTenureMoreThan + 1
        End If

    End If

Next r


End Function

'========================================================
' DEPARTMENT ANALYSIS
'========================================================

Private Sub BuildDepartmentAnalysis( _
ByVal wsData As Worksheet, _
ByVal wsAnalysis As Worksheet, _
ByVal lastRow As Long, _
ByVal outputRow As Long)


Dim departments As Object
Dim department As String
Dim r As Long
Dim item As Variant

Dim employeeCount As Long
Dim salaryTotal As Double
Dim performanceTotal As Double
Dim performanceCount As Long

Set departments = CreateObject("Scripting.Dictionary")

'--------------------------------------------
' Identify departments
'--------------------------------------------

For r = 2 To lastRow

    department = Trim(CStr(wsData.Cells(r, 3).Value))

    If department <> "" Then

        If Not departments.Exists(department) Then
            departments.Add department, department
        End If

    End If

Next r

'--------------------------------------------
' Calculate department metrics
'--------------------------------------------

For Each item In departments.Keys

    employeeCount = 0
    salaryTotal = 0
    performanceTotal = 0
    performanceCount = 0

    For r = 2 To lastRow

        If StrComp( _
            Trim(CStr(wsData.Cells(r, 3).Value)), _
            CStr(item), _
            vbTextCompare) = 0 Then

            employeeCount = employeeCount + 1

            If IsNumeric(wsData.Cells(r, 5).Value) Then
                salaryTotal = salaryTotal + _
                              CDbl(wsData.Cells(r, 5).Value)
            End If

            If IsNumeric(wsData.Cells(r, 8).Value) Then
                performanceTotal = performanceTotal + _
                                   CDbl(wsData.Cells(r, 8).Value)

                performanceCount = performanceCount + 1
            End If

        End If

    Next r

    wsAnalysis.Cells(outputRow, 4).Value = item
    wsAnalysis.Cells(outputRow, 5).Value = employeeCount

    If employeeCount > 0 Then
        wsAnalysis.Cells(outputRow, 6).Value = _
            salaryTotal / employeeCount
    End If

    If performanceCount > 0 Then
        wsAnalysis.Cells(outputRow, 7).Value = _
            performanceTotal / performanceCount
    End If

    outputRow = outputRow + 1

Next item


End Sub

'========================================================
' EMPLOYEE RISK ANALYSIS
'========================================================

Private Sub BuildRiskAnalysis( _
ByVal wsData As Worksheet, _
ByVal wsAnalysis As Worksheet, _
ByVal lastRow As Long, _
ByVal outputRow As Long)


Dim r As Long

Dim employeeID As String
Dim employeeName As String
Dim department As String

Dim salary As Double
Dim performance As Variant
Dim tenure As Double

Dim riskCategory As String
Dim recommendation As String

For r = 2 To lastRow

    employeeID = CStr(wsData.Cells(r, 1).Value)
    employeeName = CStr(wsData.Cells(r, 2).Value)
    department = CStr(wsData.Cells(r, 3).Value)

    salary = 0

    If IsNumeric(wsData.Cells(r, 5).Value) Then
        salary = CDbl(wsData.Cells(r, 5).Value)
    End If

    performance = wsData.Cells(r, 8).Value

    tenure = 0

    If IsDate(wsData.Cells(r, 6).Value) Then

        tenure = DateDiff( _
            "d", _
            CDate(wsData.Cells(r, 6).Value), _
            Date) / 365.25

    End If

    riskCategory = "Normal"
    recommendation = "Regular follow-up"

    '--------------------------------------------
    ' High performance
    '--------------------------------------------

    If IsNumeric(performance) Then

        If CDbl(performance) >= 5 Then

            riskCategory = "High Performer"
            recommendation = _
                "Consider recognition or career development"

        End If

    End If

    '--------------------------------------------
    ' Low performance
    '--------------------------------------------

    If IsNumeric(performance) Then

        If CDbl(performance) <= 2 Then

            riskCategory = "Performance Attention"
            recommendation = _
                "Consider performance follow-up"

        End If

    End If

    '--------------------------------------------
    ' Long tenure + high performance
    '--------------------------------------------

    If IsNumeric(performance) Then

        If tenure >= 5 And CDbl(performance) >= 4 Then

            riskCategory = "Key Employee"
            recommendation = _
                "Consider retention and recognition"

        End If

    End If

    '--------------------------------------------
    ' High salary + low performance
    '--------------------------------------------

    If IsNumeric(performance) Then

        If salary >= 10000 And CDbl(performance) <= 2 Then

            riskCategory = "Management Attention"
            recommendation = _
                "Review role, performance and compensation"

        End If

    End If

    '--------------------------------------------
    ' Recent hire + low performance
    '--------------------------------------------

    If IsNumeric(performance) Then

        If tenure < 1 And CDbl(performance) <= 2 Then

            riskCategory = "New Hire Attention"
            recommendation = _
                "Consider onboarding and development follow-up"

        End If

    End If

    '--------------------------------------------
    ' Write result
    '--------------------------------------------

    wsAnalysis.Cells(outputRow, 1).Value = employeeID
    wsAnalysis.Cells(outputRow, 2).Value = employeeName
    wsAnalysis.Cells(outputRow, 3).Value = department
    wsAnalysis.Cells(outputRow, 4).Value = salary
    wsAnalysis.Cells(outputRow, 5).Value = performance
    wsAnalysis.Cells(outputRow, 6).Value = tenure
    wsAnalysis.Cells(outputRow, 7).Value = riskCategory
    wsAnalysis.Cells(outputRow, 8).Value = recommendation

    outputRow = outputRow + 1

Next r


End Sub

'========================================================
' FORMAT HR ANALYSIS
'========================================================

Private Sub FormatHRAnalysis( _
ByVal ws As Worksheet, _
ByVal rowTenure As Long, _
ByVal rowDepartment As Long, _
ByVal rowRisk As Long, _
ByVal lastDepartmentRow As Long, _
ByVal lastRiskRow As Long)


Dim purple As Long
Dim green As Long
Dim lightPurple As Long
Dim lightGreen As Long
Dim darkText As Long
Dim white As Long

Dim r As Long
Dim risk As String

purple = RGB(91, 44, 131)
green = RGB(46, 125, 50)
lightPurple = RGB(238, 230, 245)
lightGreen = RGB(226, 239, 218)
darkText = RGB(45, 45, 45)
white = RGB(255, 255, 255)

'====================================================
' GENERAL
'====================================================

With ws.Cells

    .Font.Name = "Aptos"
    .Font.Size = 10
    .Font.Color = darkText

End With

'====================================================
' TITLE
'====================================================

With ws.Range("A1:J1")

    .Interior.Color = purple
    .Font.Color = white
    .Font.Bold = True
    .Font.Size = 18
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter

End With

ws.Rows(1).RowHeight = 35

With ws.Range("A2:J2")

    .Interior.Color = lightPurple
    .Font.Color = purple
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter

End With

ws.Rows(2).RowHeight = 25

'====================================================
' SECTION HEADERS
'====================================================

With ws.Range("A4:B4")

    .Interior.Color = green
    .Font.Color = white
    .Font.Bold = True
    .HorizontalAlignment = xlCenter

End With

With ws.Range("D4:E4")

    .Interior.Color = green
    .Font.Color = white
    .Font.Bold = True
    .HorizontalAlignment = xlCenter

End With

With ws.Range("G4:H4")

    .Interior.Color = green
    .Font.Color = white
    .Font.Bold = True
    .HorizontalAlignment = xlCenter

End With

With ws.Range("A" & rowTenure & ":B" & rowTenure)

    .Interior.Color = purple
    .Font.Color = white
    .Font.Bold = True
    .HorizontalAlignment = xlCenter

End With

With ws.Range( _
    "D" & rowDepartment & ":G" & rowDepartment)

    .Interior.Color = purple
    .Font.Color = white
    .Font.Bold = True
    .HorizontalAlignment = xlCenter

End With

With ws.Range("A" & rowRisk & ":H" & rowRisk)

    .Interior.Color = purple
    .Font.Color = white
    .Font.Bold = True
    .HorizontalAlignment = xlCenter

End With

'====================================================
' TABLE HEADERS
'====================================================

With ws.Range("A5:B5")

    .Interior.Color = lightPurple
    .Font.Bold = True
    .HorizontalAlignment = xlCenter

End With

With ws.Range("D5:E5")

    .Interior.Color = lightPurple
    .Font.Bold = True
    .HorizontalAlignment = xlCenter

End With

With ws.Range("G5:H5")

    .Interior.Color = lightPurple
    .Font.Bold = True
    .HorizontalAlignment = xlCenter

End With

With ws.Range( _
    "A" & rowTenure + 1 & ":B" & rowTenure + 1)

    .Interior.Color = lightPurple
    .Font.Bold = True
    .HorizontalAlignment = xlCenter

End With

With ws.Range( _
    "D" & rowDepartment + 1 & ":G" & rowDepartment + 1)

    .Interior.Color = lightPurple
    .Font.Bold = True
    .HorizontalAlignment = xlCenter

End With

With ws.Range( _
    "A" & rowRisk + 1 & ":H" & rowRisk + 1)

    .Interior.Color = lightPurple
    .Font.Bold = True
    .HorizontalAlignment = xlCenter

End With

'====================================================
' NUMBER FORMATS
'====================================================

' Currency
ws.Range("E6:E10").NumberFormat = _
    "[$R$-pt-BR] #,##0.00"

ws.Range( _
    "F" & rowDepartment + 2 & ":F" & lastDepartmentRow).NumberFormat = _
    "[$R$-pt-BR] #,##0.00"

ws.Range( _
    "D" & rowRisk + 2 & ":D" & lastRiskRow).NumberFormat = _
    "[$R$-pt-BR] #,##0.00"

' Average performance
ws.Range("H6").NumberFormat = "0.00"

ws.Range( _
    "G" & rowDepartment + 2 & ":G" & lastDepartmentRow).NumberFormat = _
    "0.00"

' Tenure in years
ws.Range( _
    "B" & rowTenure + 2 & ":B" & rowTenure + 3).NumberFormat = _
    "0.00"

' Employee quantities
ws.Range( _
    "B" & rowTenure + 4 & ":B" & rowTenure + 7).NumberFormat = _
    "0"

' Risk analysis
ws.Range( _
    "F" & rowRisk + 2 & ":F" & lastRiskRow).NumberFormat = _
    "0.00"

'====================================================
' BORDERS — ONLY AREAS WITH DATA
'====================================================

With ws.Range("A5:B9").Borders

    .LineStyle = xlContinuous
    .Color = RGB(215, 215, 215)
    .Weight = xlThin

End With

With ws.Range("D5:E10").Borders

    .LineStyle = xlContinuous
    .Color = RGB(215, 215, 215)
    .Weight = xlThin

End With

With ws.Range("G5:H10").Borders

    .LineStyle = xlContinuous
    .Color = RGB(215, 215, 215)
    .Weight = xlThin

End With

With ws.Range( _
    "A" & rowTenure + 1 & ":B" & rowTenure + 7).Borders

    .LineStyle = xlContinuous
    .Color = RGB(215, 215, 215)
    .Weight = xlThin

End With

If lastDepartmentRow >= rowDepartment + 2 Then

    With ws.Range( _
        "D" & rowDepartment + 1 & ":G" & lastDepartmentRow).Borders

        .LineStyle = xlContinuous
        .Color = RGB(215, 215, 215)
        .Weight = xlThin

    End With

End If

If lastRiskRow >= rowRisk + 2 Then

    With ws.Range( _
        "A" & rowRisk + 1 & ":H" & lastRiskRow).Borders

        .LineStyle = xlContinuous
        .Color = RGB(215, 215, 215)
        .Weight = xlThin

    End With

End If

'====================================================
' RISK COLORS
'====================================================

For r = rowRisk + 2 To lastRiskRow

    risk = LCase(Trim(CStr(ws.Cells(r, 7).Value)))

    Select Case risk

        Case "high performer", "key employee"

            ws.Cells(r, 7).Interior.Color = lightGreen
            ws.Cells(r, 7).Font.Color = green
            ws.Cells(r, 7).Font.Bold = True

        Case "performance attention", _
             "new hire attention"

            ws.Cells(r, 7).Interior.Color = RGB(255, 243, 205)
            ws.Cells(r, 7).Font.Color = RGB(156, 101, 0)
            ws.Cells(r, 7).Font.Bold = True

        Case "management attention"

            ws.Cells(r, 7).Interior.Color = RGB(242, 222, 222)
            ws.Cells(r, 7).Font.Color = RGB(156, 0, 0)
            ws.Cells(r, 7).Font.Bold = True

    End Select

Next r

'====================================================
' ALIGNMENT
'====================================================

ws.Range("B6:B9").HorizontalAlignment = xlCenter

ws.Range("E6:E10").HorizontalAlignment = xlRight

ws.Range("H6:H10").HorizontalAlignment = xlCenter

ws.Range( _
    "B" & rowTenure + 2 & ":B" & rowTenure + 7).HorizontalAlignment = xlCenter

ws.Range( _
    "E" & rowDepartment + 2 & ":G" & lastDepartmentRow).HorizontalAlignment = xlCenter

ws.Range( _
    "D" & rowRisk + 2 & ":F" & lastRiskRow).HorizontalAlignment = xlCenter

'====================================================
' COLUMN WIDTHS
'====================================================

ws.Columns("A").ColumnWidth = 18
ws.Columns("B").ColumnWidth = 27
ws.Columns("C").ColumnWidth = 22
ws.Columns("D").ColumnWidth = 25
ws.Columns("E").ColumnWidth = 18
ws.Columns("F").ColumnWidth = 14
ws.Columns("G").ColumnWidth = 24
ws.Columns("H").ColumnWidth = 45
ws.Columns("I:J").ColumnWidth = 12

'====================================================
' ROW HEIGHT
'====================================================

ws.Rows("4:" & lastRiskRow).RowHeight = 20

'====================================================
' FILTER RISK ANALYSIS
'====================================================

If ws.AutoFilterMode Then
    ws.AutoFilterMode = False
End If

If lastRiskRow >= rowRisk + 2 Then

    ws.Range( _
        "A" & rowRisk + 1 & ":H" & lastRiskRow).AutoFilter

End If

'====================================================
' FREEZE PANES
' DISABLED
'====================================================

ws.Activate

ActiveWindow.FreezePanes = False
ActiveWindow.SplitRow = 0
ActiveWindow.SplitColumn = 0

End Sub


