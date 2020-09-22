Attribute VB_Name = "mdlStandard"
Option Explicit

Public gStrConnectionString As String
Public gIntUserId As Integer
Public gStrUser As String 'store user code

Public gIntDepartment As Integer 'modification/deletion
Public gBlDepartment As Boolean

Public gIntDesignation As Integer 'modification/deletion
Public gBlDesignation As Boolean

Public gIntSubDepartment As Integer 'modification/deletion
Public gBlSubDepartment As Boolean

Public gIntHoliday As Integer 'modification/deletion
Public gBlHoliday As Boolean

Public gIntShift As Integer 'modification/deletion
Public gBlShift As Boolean

Public gIntProduct As Integer 'modification/deletion
Public gBlProduct As Boolean

Public gIntVehicle As Integer 'modification/deletion
Public gBlVehicle As Boolean

Public gIntGuestHOuse As Integer 'modification/deletion
Public gBlGuesthouse As Boolean

Public gIntEmployee As Integer 'modification/deletion
Public gBlEmployee As Boolean

Public gIntGuesthouseBooking As Integer 'modification/deletion
Public gBlGuesthouseBooking As Boolean

Public gIntIssuePass As Integer 'modification/deletion
Public gBlIssuePass As Boolean

Public gIntVehicleBooking As Integer 'modification/deletion
Public gBlVehicleBooking As Boolean

Public gIntVehicleHire As Integer 'modification/deletion
Public gBlVehicleHire As Boolean

Public gIntRequirement As Integer 'modification/deletion
Public gBlRequirement As Boolean

Public gIntPaymentOfBill As Integer 'modification/deletion
Public gBlPaymentOfBill As Boolean

Public gIntVehicleUse As Integer 'modification/deletion
Public gBlVehicleUse As Boolean

Public gIntLeave As Integer 'modification/deletion
Public gBlLeave As Boolean

Public gIntCanteenItemType As Integer 'modification/deletion
Public gBlCanteenItemType As Boolean

Public gIntCanteenItem As Integer 'modification/deletion
Public gBlCanteenItem As Boolean

Public gBlListDept As Boolean 'get list
Public gstrListDeptName As String
Public gintListDeptId As Integer

Public gBlListSubDept As Boolean
Public gstrListSubDeptName As String
Public gintListSubDeptId As Integer

Public gBlListDesg As Boolean
Public gstrListDesgName As String
Public gintListDesgId As Integer

Public gBlListShift As Boolean
Public gstrListShiftName As String
Public gintListShiftId As Integer

Public gBlListEmployee As Boolean
Public gstrListEmployee As String
Public gintListEmployee As Integer

Public gBlListGuesthouse As Boolean
Public gstrListGuesthouse As String
Public gintListGuesthouse As Integer

Public gBlListRoom As Boolean
Public gstrListRoom As String
Public gintListRoom As Integer

Public gBlListAvailableVehicle As Boolean
Public gstrListAvailableVehicle As String
Public gintListAvailableVehicle As Integer

Public gBlListOrder As Boolean
Public gstrListOrder As String
Public gintListOrder As Integer

Public gBlListVehicleNum As Boolean
Public gstrListVehicleNum As String
Public gintListVehicleNum As Integer

Public gBlListCanteenItemType As Boolean
Public gstrListCanteenItemType As String
Public gintListCanteenItemType As Integer

Public gBlListApproveBy As Boolean
Public gstrListApproveBy As String
Public gintListApproveBy As Integer

Public gBlPayslipGenComplete As Boolean

Public Sub gFormCenter(frm As Form)
    frm.Left = (Screen.Width - frm.Width) / 2
    frm.Top = (Screen.Height - frm.Height) / 2
End Sub

Public Sub gCalculateSalary()
On Error GoTo SalaryCalculationError

Dim cnCalc As New ADODB.Connection, strSql As String, strDateCrt As String
Dim rsCalc As New ADODB.Recordset, i As Integer, dDate As Date, iYearValue As Integer, iMonthValue

iMonthValue = Month(Date)
    If iMonthValue <> 1 Then
        strDateCrt = "#1/" & iMonthValue - 1 & "/" & Year(Date) & "# and #" & _
            gCalcDays(Month(Date) - 1, Year(Date)) & "/" & iMonthValue - 1 & "/" & Year(Date) & "#"
            dDate = "1/" & Month(Date) - 1 & "/" & Year(Date)
            iYearValue = Year(Date)
    ElseIf iMonthValue = 1 Then
        strDateCrt = "#1/dec/" & Year(Date) - 1 & "# and 31/dec/" & Year(Date) - 1 & "#"
            dDate = "1/dec/" & Year(Date) - 1
            iYearValue = Year(Date) - 1
    End If
    'update present days
    strSql = "Select emp_code,count(*) from t_employee_attendence where  ea_date between " & strDateCrt & " group by emp_code "
            
    rsCalc.Open strSql, gStrConnectionString, adOpenKeyset, adLockOptimistic
    
    If rsCalc.RecordCount > 0 Then
        cnCalc.ConnectionString = gStrConnectionString
        cnCalc.Open
        cnCalc.BeginTrans
        For i = 0 To rsCalc.RecordCount - 1
            strSql = "Insert into t_monthly_attendence(ma_month,emp_code,ma_present) values ('" & _
                Format(DateAdd("m", -1, Date), "MMM-yyyy") & "','" & rsCalc.Fields(0) & "'," & _
                rsCalc.Fields(1) & ")"
            cnCalc.Execute strSql
            rsCalc.MoveNext
        Next
        DoEvents
        cnCalc.CommitTrans
    End If
    'update holidays including sundays
    If rsCalc.State = 1 Then rsCalc.Close
        
    rsCalc.Open "select count(*) from m_holidays where hd_date between " & strDateCrt _
            , gStrConnectionString, adOpenKeyset, adLockOptimistic
    
    Dim iCntSunday As Integer

    For i = 1 To gCalcDays(Month(Date) - 1, iYearValue)
        If Weekday(dDate) = 1 Then
            iCntSunday = iCntSunday + 1
        End If
        dDate = DateAdd("d", 1, dDate)
    Next
    DoEvents
    
    strSql = "Update t_monthly_attendence set ma_holiday=" & iCntSunday + Val(rsCalc.Fields(0)) & _
            " where ma_month= '" & Format(DateAdd("m", -1, Date), "MMM-yyyy") & "'"
    cnCalc.Execute strSql
    'update leave
    Dim rsList As New ADODB.Recordset, j As Integer, dtTempDate As Date
    Dim iCntSun As Integer, iCntLeave As Integer, iCntHoliday As Integer
    
    If rsCalc.State = 1 Then rsCalc.Close
        
    rsList.Open " select distinct tea.emp_code,me.emp_id from t_employee_attendence tea,m_employee me where tea.emp_code <> '0' and tea.emp_code=me.emp_code", gStrConnectionString, adOpenKeyset, adLockOptimistic
    
    For i = 0 To rsList.RecordCount - 1
        If rsCalc.State = 1 Then rsCalc.Close
        strSql = "select lr_from ,lr_to from t_leave_registration where emp_id=" & rsList.Fields(1) & " and status=1"
        rsCalc.Open strSql, gStrConnectionString, adOpenKeyset, adLockOptimistic
        If rsCalc.RecordCount > 0 Then
            dtTempDate = rsCalc.Fields(0)
            For j = 1 To DateDiff("d", rsCalc.Fields(0), rsCalc.Fields(1))
                If Month(dtTempDate) = Month(Date) - 1 Then
                    iCntLeave = iCntLeave + 1
                Else
                    Exit For
                End If
                If Weekday(dtTempDate) = 1 Then
                    iCntSun = iCntSun + 1
                Else
                    Dim rsTemp As New ADODB.Recordset
                    Set rsTemp = cnCalc.Execute("select * from m_holidays where hd_date=#" & dtTempDate & "#")
                    If rsTemp.RecordCount > 0 Then
                        iCntHoliday = iCntHoliday + 1
                    End If
                End If
                dtTempDate = DateAdd("d", 1, dtTempDate)
            Next
        End If
        
        strSql = "Update t_monthly_attendence set ma_leave=" & iCntLeave & ",ma_holiday=ma_holiday - " & _
                (iCntHoliday + iCntSun) & " where emp_code='" & rsList.Fields(0) & "' and " & _
                " ma_month ='" & Format(DateAdd("m", -1, Date), "MMM-yyyy") & "'"
        cnCalc.Execute strSql
        DoEvents
        rsList.MoveNext
    Next
    'calculate final salary
    Dim cSalary As Currency, cPF As Currency
    
    If rsCalc.State = 1 Then rsCalc.Close
    rsCalc.Open "select * from view_Salarycalc", gStrConnectionString, adOpenKeyset, adLockOptimistic
    
    cnCalc.BeginTrans
    For i = 0 To rsCalc.RecordCount - 1
        cPF = 0
        cSalary = rsCalc.Fields("emp_basic") / gCalcDays(IIf(iMonthValue = 1, 12, iMonthValue - 1), iYearValue) _
            * (rsCalc.Fields("ma_present") + rsCalc.Fields("ma_holiday") + rsCalc.Fields("ma_leave"))
        If rsCalc.Fields("emp_pf") = "Yes" Then
            cPF = cSalary * 0.12
        End If
        cnCalc.Execute "update t_monthly_attendence set ma_salary=" & cSalary & ",ma_pf=" & cPF & _
            " where emp_code='" & rsCalc.Fields("emp_code") & "' and ma_month='" & _
            Format(DateAdd("m", -1, Date), "MMM-yyyy") & "'"
        rsCalc.MoveNext
    Next
    cnCalc.CommitTrans
    gBlPayslipGenComplete = True
    MsgBox "Salary calculation is complete", vbInformation, "Salary Calculation"
    DoEvents
Exit Sub
SalaryCalculationError:
    MsgBox "Error :" & Err.Number & vbCrLf & "Desc :" & Err.Description, vbCritical, "Salary Calculation Error"
    Err.Clear
End Sub

Public Function gCalcDays(iMonth As Integer, iYear As Integer) As Integer
    Select Case iMonth
        Case 1, 3, 5, 7, 8, 10, 12
            gCalcDays = 31
        Case 4, 6, 9, 11
            gCalcDays = 30
        Case 2
            If iYear Mod 4 = 0 Then
                gCalcDays = 29
            Else
                gCalcDays = 28
            End If
    End Select
End Function

Public Function gCheckSalaryCalculation(Optional strMonth As String) As Boolean
On Error GoTo SalCheckError
    Dim rsCheck As New ADODB.Recordset, strQ As String
'    If gBlPayslipGenComplete = True Then
'        gCheckSalaryCalculation = True
'    End If
    
    If Len(strMonth) = 0 Then
        strQ = "select * from t_monthly_attendence where ma_month='" & Format(DateAdd("m", -1, Date), "MMM-yyyy") & "'"
    Else
        strQ = "select * from t_monthly_attendence where ma_month='" & strMonth & "'"
    End If
    
    rsCheck.Open strQ, gStrConnectionString, adOpenKeyset, adLockOptimistic
    If rsCheck.RecordCount > 0 Then
        gCheckSalaryCalculation = True
    Else
        gCheckSalaryCalculation = False
    End If
Exit Function
SalCheckError:
    MsgBox "Error :" & Err.Number & vbCrLf & "Desc :" & Err.Description, vbCritical, "Salary Check Error"
    Err.Clear
End Function
