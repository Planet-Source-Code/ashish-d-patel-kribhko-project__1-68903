VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Kribhko Management System"
   ClientHeight    =   6060
   ClientLeft      =   2025
   ClientTop       =   2145
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6060
   ScaleWidth      =   6840
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2670
      Top             =   2670
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7185
      Left            =   3870
      TabIndex        =   7
      Top             =   450
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   12674
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton cmdReports 
      Caption         =   "&Reports"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   5
      Top             =   7245
      Width           =   3735
   End
   Begin VB.CommandButton cmdOthers 
      Caption         =   "&Others"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   4
      Top             =   6870
      Width           =   3735
   End
   Begin VB.CommandButton cmdPersonnelInfo 
      Caption         =   "Personnel &Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   3
      Top             =   6495
      Width           =   3735
   End
   Begin VB.CommandButton cmdFinance 
      Caption         =   "&Finance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   2
      Top             =   6120
      Width           =   3735
   End
   Begin VB.CommandButton cmdPurchase 
      Caption         =   "&Purchase System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   1
      Top             =   5745
      Width           =   3735
   End
   Begin VB.CommandButton cmdManagement 
      Caption         =   "Personnel Management && &Admin."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   0
      Top             =   5370
      Width           =   3735
   End
   Begin MSComctlLib.TreeView trvList 
      Height          =   5265
      Left            =   90
      TabIndex        =   6
      Top             =   60
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   9287
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3870
      TabIndex        =   8
      Top             =   90
      Width           =   7875
   End
   Begin VB.Menu mnuright 
      Caption         =   "Right"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add or Modify"
      End
   End
   Begin VB.Menu mnuPrintit 
      Caption         =   "Printit"
      Visible         =   0   'False
      Begin VB.Menu mnuView 
         Caption         =   "View"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strCaption As String
Private Sub cmdFinance_Click()
    
    With trvList.Nodes
        .Clear
        .Add , , "a", "Finance"
        .Add "a", tvwChild, , "Payment of Bills"
    End With
    
    trvList.Nodes(1).Expanded = True
    
    cmdFinance.Enabled = False
    
    cmdManagement.Enabled = True
    cmdOthers.Enabled = True
    cmdPersonnelInfo.Enabled = True
    cmdPurchase.Enabled = True
    cmdReports.Enabled = True
    
End Sub

Private Sub cmdManagement_Click()

    With trvList.Nodes
        .Clear
        .Add , , "a", "Personnel Management & Admin."
        .Add "a", tvwChild, "b", "Master Entry"
        .Add "b", tvwChild, , "Department"
        .Add "b", tvwChild, , "Sub Department"
        .Add "b", tvwChild, , "Designation"
        .Add "b", tvwChild, , "Vehicle"
        .Add "b", tvwChild, , "Guesthouse"
        .Add "b", tvwChild, , "Shift"
        .Add "b", tvwChild, , "Product Purchase"
        .Add "b", tvwChild, , "Holiday"
        .Add "b", tvwChild, , "Employee"
        .Add "a", tvwChild, , "Guesthouse Booking"
        .Add "a", tvwChild, , "Issue Bus Pass"
        .Add "a", tvwChild, , "Adv. Booking for Vehicle"
        .Add "a", tvwChild, , "Employee Leave Saction"
        .Add "a", tvwChild, , "Employee Shift Sch."
        '.Add "a", tvwChild, , "Pay Slip Generate"
    End With
    
    trvList.Nodes(1).Expanded = True
    trvList.Nodes(1).Child.Expanded = True
    
    cmdManagement.Enabled = False
    
    cmdPurchase.Enabled = True
    cmdFinance.Enabled = True
    cmdOthers.Enabled = True
    cmdReports.Enabled = True
    cmdPersonnelInfo.Enabled = True
End Sub

Private Sub cmdOthers_Click()
    
    With trvList.Nodes
        .Clear
        .Add , , "a", "Others"
        .Add "a", tvwChild, "b", "Canteen"
        .Add "b", tvwChild, , "Item Type"
        .Add "b", tvwChild, , "Item"
        .Add "b", tvwChild, , "New Bill"
        .Add "b", tvwChild, , "Duplicate Bill"
        .Add "a", tvwChild, "c", "Search Employee"
    End With
    
    trvList.Nodes(1).Expanded = True
    trvList.Nodes(1).Child.Expanded = True
    cmdOthers.Enabled = False
    
    cmdFinance.Enabled = True
    cmdManagement.Enabled = True
    cmdPersonnelInfo.Enabled = True
    cmdPurchase.Enabled = True
    cmdReports.Enabled = True
End Sub

Private Sub cmdPersonnelInfo_Click()
    
    With trvList.Nodes
        .Clear
        .Add , , "a", "Personnel Information"
        .Add "a", tvwChild, , "Change Password"
        .Add "a", tvwChild, , "Vehicle Usage"
        .Add "a", tvwChild, , "Employee Info"
        '.Add "a", tvwChild, , "Leave Entry"
        .Add "a", tvwChild, , "Time Card"
        .Add "a", tvwChild, , "Monthly Shift Schedule"
        '.Add "a", tvwChild, , "Advance Salary"
        .Add "a", tvwChild, , "Payslip"
        
    End With
    
    trvList.Nodes(1).Expanded = True
    
    cmdPersonnelInfo.Enabled = False
    If gIntUserId = 0 Then
        cmdFinance.Enabled = True
        cmdManagement.Enabled = True
        cmdOthers.Enabled = True
        cmdPurchase.Enabled = True
        cmdReports.Enabled = True
    End If
End Sub

Private Sub cmdPurchase_Click()
    
    With trvList.Nodes
        .Clear
        .Add , , "a", "Purchase System"
        .Add "a", tvwChild, , "Vehicle Hire"
        .Add "a", tvwChild, , "Requirement Entry"
        .Add "a", tvwChild, , "Receive Ordered Product"
    End With
    
    trvList.Nodes(1).Expanded = True
    
    cmdPurchase.Enabled = False
    
    cmdManagement.Enabled = True
    cmdFinance.Enabled = True
    cmdOthers.Enabled = True
    cmdReports.Enabled = True
    cmdPersonnelInfo.Enabled = True
End Sub

Private Sub cmdReports_Click()
    
    With trvList.Nodes
        .Clear
        .Add , , "a", "Reports"
        .Add "a", tvwChild, "b", "Master List"
        
        .Add "b", tvwChild, , "Department"
        .Add "b", tvwChild, , "Subdepartments"
        .Add "b", tvwChild, , "Designation"
        .Add "b", tvwChild, , "Shift"
        .Add "b", tvwChild, , "Hired Vehicle"
        .Add "b", tvwChild, , "Product"
        .Add "b", tvwChild, , "Employee"
        .Add "b", tvwChild, , "Holidays"
        
        .Add "a", tvwChild, "c", "Canteen"
        .Add "c", tvwChild, , "Departmentwise Expense"
        .Add "c", tvwChild, , "Employeewise Expense"
        
        .Add "a", tvwChild, "d", "Purchase"
        .Add "d", tvwChild, , "Duplicate Copy of Order"
        .Add "d", tvwChild, , "Purchase By Department"
        '.Add "c", tvwChild, "Datewise Expense"
        
'        .Add "a", tvwChild, "c", "Employee By"
'        .Add "c", tvwChild, , "Department"
'        .Add "c", tvwChild, , "Designation"
'        .Add "c", tvwChild, , "Shift"
'        .Add "c", tvwChild, , "Bloodgroup"
'        .Add "c", tvwChild, , "Basic Salary"
'        .Add "c", tvwChild, , "Birthdate"
'        .Add "c", tvwChild, , "Date of Join"
'
'        .Add "a", tvwChild, , "Guesthouse Booking"
'        .Add "a", tvwChild, , "Issued Buspass"
'        .Add "a", tvwChild, , "Vehicle Booking"
'        .Add "a", tvwChild, , "Attendence Register"
'        .Add "a", tvwChild, , "Leave Register"
'        .Add "a", tvwChild, , "Payslip"
'
'        .Add "a", tvwChild, "d", "Labels"
'        .Add "d", tvwChild, , "Employee"
'        .Add "d", tvwChild, , "Vehicle owner"
    End With
    
    trvList.Nodes(1).Expanded = True
    trvList.Nodes(1).Child.Expanded = True
    trvList.Nodes(11).Expanded = True
    trvList.Nodes(14).Expanded = True
    
    cmdReports.Enabled = False
    
    cmdFinance.Enabled = True
    cmdManagement.Enabled = True
    cmdOthers.Enabled = True
    cmdPersonnelInfo.Enabled = True
    cmdPurchase.Enabled = True
End Sub

Private Sub Form_Load()
On Error GoTo MainLoadError
    Dim strStat As String, strStat2 As String, dtDate As Date
    Dim rsTest As New ADODB.Recordset, rsTest2 As New ADODB.Recordset
    dtDate = DateAdd("m", -1, Date)
    strStat = "select * from t_employee_attendence where month(ea_date) =" & Month(dtDate) & _
                " and year(ea_date)=" & Year(dtDate)
    strStat2 = "select * from t_monthly_attendence where ma_month='" & Format(DateAdd("m", -1, Date), "MMM-yyyy") & "'"
    rsTest.Open strStat, gStrConnectionString, adOpenKeyset, adLockOptimistic
    rsTest2.Open strStat2, gStrConnectionString, adOpenKeyset, adLockOptimistic
    If rsTest.RecordCount > 0 And rsTest2.RecordCount = 0 Then
        frmGeneratePayslip.strMonth = Format(DateAdd("m", -1, Date), "MMM-yyyy")
        frmGeneratePayslip.Show vbModal
    End If
    rsTest.Close
    rsTest2.Close
    
    If gIntUserId <> 0 Then
        cmdFinance.Enabled = False
        cmdManagement.Enabled = False
        cmdOthers.Enabled = False
        cmdPurchase.Enabled = False
        cmdReports.Enabled = False
    End If
    
Exit Sub
MainLoadError:
    MsgBox "Error :" & Err.Number & vbCrLf & "Desc : " & Err.Description, vbCritical, "Main Load Error"
    Err.Clear
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 2 And strCaption = "Duplicate Bill" Then
    PopupMenu mnuPrintit
End If
End Sub

Private Sub mnuAdd_Click()
    Call mShowForm(strCaption)
End Sub

Private Sub mnuView_Click()
    If Len(ListView1.ListItems(1).Text) > 0 Then
        CrystalReport1.DataFiles(0) = App.Path & "\KSystemdb.mdb"

        If ListView1.SelectedItem.ListSubItems(2).Text = "ATR" Then
            CrystalReport1.ReportFileName = App.Path & "\Report\rptDepartmentCanteenBill.rpt"
            CrystalReport1.ReplaceSelectionFormula "{Rpt_CanteenBillForDepartment.b_no}=" & ListView1.SelectedItem.Text
        Else
            CrystalReport1.ReportFileName = App.Path & "\Report\rptEmployeeCanteenBill.rpt"
            CrystalReport1.ReplaceSelectionFormula "{rpt_EmployeeCanteenBill.b_no}=" & ListView1.SelectedItem.Text
        End If

        CrystalReport1.PrintReport
        CrystalReport1.WindowState = crptMaximized
    End If
End Sub

Private Sub trvList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And strCaption <> "Duplicate Bill" Then
        PopupMenu mnuRight
    End If
End Sub

Private Sub trvList_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo ReportError
Dim strSql As String
'    If Node.Parent = "Nothing" Then Exit Sub
    'If Node.Parent = "Reports" Or Node.Parent = "Master List" Or Node.Parent = "Employee By" Or Node.Parent = "Labels" Then
     If Node.Parent = "Reports" Or Node.Parent = "Master List" Or Node.Parent = "Canteen" Or Node.Parent = "Purchase" Then
        ListView1.ListItems.Clear
        ListView1.ColumnHeaders.Clear
        Label1.Caption = ""
        
        With CrystalReport1
            .DataFiles(0) = App.Path & "\Ksystemdb.mdb"
            .WindowState = crptMaximized
            If Node.Parent = "Master List" Then
                If Node.Text = "Department" Then
                    .ReportFileName = App.Path & "\Report\mlrptDepartment.rpt"
                    .ReportTitle = "Master List of Departments"
                    .PrintReport
                ElseIf Node.Text = "Subdepartments" Then
                    .ReportFileName = App.Path & "\Report\mlrptsubDepartment.rpt"
                    .ReportTitle = "Master List of Subdepartments"
                    .PrintReport
                ElseIf Node.Text = "Designation" Then
                    .ReportFileName = App.Path & "\Report\mlrptDesignation.rpt"
                    .ReportTitle = "Master List of Designation"
                    .PrintReport
                ElseIf Node.Text = "Holidays" Then
                    .ReportFileName = App.Path & "\Report\mlrptholidays.rpt"
                    .ReportTitle = "Master List of Holidays"
                    .PrintReport
                ElseIf Node.Text = "Hired Vehicle" Then
                    .ReportFileName = App.Path & "\Report\mlrpthiredvehicle.rpt"
                    .ReportTitle = "Master List of Hired Vehicle"
                    .PrintReport
                ElseIf Node.Text = "Shift" Then
                    .ReportFileName = App.Path & "\Report\mlrptshift.rpt"
                    .ReportTitle = "Master List of Shift"
                    .PrintReport
                ElseIf Node.Text = "Employee" Then
                    .ReportFileName = App.Path & "\Report\mlrptemployee.rpt"
                    .ReportTitle = "Master List of Employee"
                    .PrintReport
                ElseIf Node.Text = "Product" Then
                    .ReportFileName = App.Path & "\Report\mlrptproducts.rpt"
                    .ReportTitle = "Master List of Product from Purchase Department"
                    .PrintReport
                End If
            ElseIf Node.Parent = "Canteen" Then
                If Node.Text = "Departmentwise Expense" Then
                    .ReportFileName = App.Path & "\report\RptDepartmentwiseExpense.rpt"
                    .ReportTitle = "Departmentwise Expense in Canteen"
                    .PrintReport 'Employeewise Expense
                ElseIf Node.Text = "Employeewise Expense" Then
                    .ReportFileName = App.Path & "\report\rptemployeewiseexpense.rpt"
                    .ReportTitle = "Employeewise Expense in Canteen"
                    .PrintReport
                End If
            ElseIf Node.Parent = "Purchase" Then
                If Node.Text = "Duplicate Copy of Order" Then
                    Dim str As String
                    
                    str = InputBox("Enter Order Number", "Duplicate Copy of Order")
                    If Len(str) > 0 Then
                        .ReportFileName = App.Path & "\report\rptorderdetails.rpt"
                        .ReportTitle = "Order Details"
                        .ReplaceSelectionFormula "{rpt_OrderDetail.Re_Order_Id}='" & str & "'"
                        .PrintReport
                    End If
                ElseIf Node.Text = "Purchase By Department" Then
                    .ReportFileName = App.Path & "\report\rptPurchasebydepartment.rpt"
                    .ReportTitle = "Summary of Purchase, Departmentwise"
                    .PrintReport 'RptPurchaseByDepartment.rpt
                End If
            End If
        End With
        Exit Sub
    End If
    
    If Node.Text = "Department" Then
        strSql = "select dept_id as Num,dept_name as Department from m_department where status=1"
    ElseIf Node.Text = "Designation" Then
        strSql = "select desg_id as Num,desg_name as Designation from m_designation where status=1"
    ElseIf Node.Text = "Sub Department" Then
        strSql = "select sub_dept_id as Num, dept_name as Department,sub_dept_name as Subdepartment from view_subdepartment where status=1"
    ElseIf Node.Text = "Holiday" Then
        strSql = "select hd_id as Num,hd_date as [Date],hd_desc as Description from m_holidays where status=1"
    ElseIf Node.Text = "Shift" Then
        strSql = " select shf_id as Num,shf_name as Name,shf_start_time as [Start Time],shf_end_time as [End Time],shf_late_entry as [Late allowed],shf_early_go as [Early go],shf_remark as Remark from m_shift where status=1"
    ElseIf Node.Text = "Product Purchase" Then
        strSql = "select prt_id as Num,prt_name as Product,prt_comp as Company,prt_product_type as [Type],prt_price as Price from m_product where status=1"
    ElseIf Node.Text = "Vehicle" Then
        strSql = "select veh_id as Num,veh_make_comp as Company,veh_seat_available as [Seat available],veh_type as Type,veh_fuel as Fuel from m_vehicle where status=1"
    ElseIf Node.Text = "Guesthouse" Then
        strSql = "select gh_id as Num,gh_location as Location,gh_remark as Remark from m_guesthouse where status=1"
    ElseIf Node.Text = "Employee" Then
        strSql = "select emp_id as ID,emp_code as [Emp Code],emp_fname as Firstname,emp_lname as [Last Name],emp_join_date as [Join Date] from m_employee where status=1"
    ElseIf Node.Text = "Guesthouse Booking" Then
        strSql = "select ghb_id as Num, ghb_date as [Date],gh_location as [Location],ghd_room_no as [Room No],emp_fname as [Employee] from view_guesthousebooking where status=1"
    ElseIf Node.Text = "Issue Bus Pass" Then
        strSql = "select ibp_id as Num,ibp_date as [Issue Date],ibp_serial_no as [Serial No],emp_fname as [Employee],ibp_Type as [Type] from view_issuebuspass where status=1"
    ElseIf Node.Text = "Adv. Booking for Vehicle" Then
        strSql = "select vhb_id as Num,vhb_date as [Date],emp_fname as Employee,vhb_alloted_from as [From Date],vhb_alloted_to as [To Date] from view_vehiclebooking where status=1"
    ElseIf Node.Text = "Employee Shift Sch." Then
        strSql = "select ess_id as Num,emp_fname as Employee,shf_name as Shift,ess_from as [Start Date],ess_to as [End Date] from view_shiftschedule where status=1"
    ElseIf Node.Text = "Vehicle Hire" Then
        strSql = "select vhh_id as Num,veh_make_comp as Vehicle,vhh_hired_date_from as [From Date],vhh_hired_date_to as [To Date] from view_vehiclehire where status=1"
    ElseIf Node.Text = "Requirement Entry" Then
        strSql = "select re_id as Num,re_order_id as [OrderNum],dept_name as Department,re_date as [Req Date] from view_requirement where status=1"
    ElseIf Node.Text = "Receive Ordered Product" Then
        strSql = "select re_id as Num,re_order_id as [OrderNum],re_order_recv_date as [Receive Date] from view_orderreceive where status=1"
    ElseIf Node.Text = "Payment of Bills" Then
        strSql = "select pb_id as Num,pb_date as [Date],emp_fname as Employee,pb_billno as [BillNum],pb_about as [Type]  from view_paymentofbills where status=1"
    ElseIf Node.Text = "Vehicle Usage" Then
        strSql = "select vhu_id as Num,vhu_date as [Date], vhh_number_plate as Vehicle,vhu_start_km as [Start Km],vhu_end_km as [End Km] from view_vehicleuse where status=1"
    ElseIf Node.Text = "Employee Leave Saction" Then
        strSql = "select lr_id as Num,lr_date as [Date],Employee,lr_from as [From Date],lr_to as [To Date],lr_reason as Reason from view_leaveregistration where status=1"
    ElseIf Node.Text = "Item Type" Then
        strSql = "select it_id as Num,it_Name as [Item Type] from m_canteen_itemtype where status=1"
    ElseIf Node.Text = "Item" Then
        strSql = "select i_id as Num,i_name as Item from m_canteen_item where status=1"
    ElseIf Node.Text = "Duplicate Bill" Then
        strSql = "select b_no as Num,b_date as [Bill Date],b_category as Category,emp_fname as [Approved By] from view_canteenbill where status=1"
    ElseIf Node.Text = "Monthly Shift Schedule" Then
        strSql = "select ess_id as Num,emp_fname as Employee,shf_name as Shift ,ess_from as [Start Date],ess_to as [End Date] from view_monthlyshiftschedule where status=1"
    ElseIf Node.Text = "Time Card" Then
        strSql = "select ea_id as Num,ea_date as [Date],emp_fname as Employee,ea_in_time as [In Time],ea_out_time as [Out Time] from view_timecard"
    ElseIf Node.Text = "Search Employee" Then
        frmSearchEmployee.Show vbModal
    End If
    
    
    
    Label1.Caption = Node.Text
    strCaption = Label1.Caption
    If Len(strSql) = 0 Then Exit Sub
    Call mFillListView(strSql)
    Exit Sub
ReportError:
Err.Clear
End Sub

Public Sub mFillListView(strSql As String)

    Dim lstItem As ListItem, i As Integer, j As Integer
    Dim rsListView As New ADODB.Recordset
    
    ListView1.ColumnHeaders.Clear
    ListView1.ListItems.Clear
    
    rsListView.Open strSql, gStrConnectionString, adOpenKeyset, adLockOptimistic
    
    For i = 1 To rsListView.Fields.Count
        ListView1.ColumnHeaders.Add , , rsListView.Fields(i - 1).Name
    Next
    
    For i = 0 To rsListView.RecordCount - 1
        Set lstItem = ListView1.ListItems.Add(, , rsListView.Fields(0))
        For j = 1 To rsListView.Fields.Count - 1
            lstItem.SubItems(j) = rsListView.Fields(j)
        Next
        rsListView.MoveNext
    Next
    
    ListView1.ColumnHeaders(1).Width = 700
    For i = 1 To rsListView.Fields.Count - 1
        ListView1.ColumnHeaders(i + 1).Width = (ListView1.Width - 1000) / (rsListView.Fields.Count - 1)
    Next
    
End Sub

Public Sub mShowForm(strFormName As String)
    
    If strFormName = "Department" Then
        frmDepartmentMaster.Show vbModal
    ElseIf strFormName = "Designation" Then
        frmDesignationMaster.Show vbModal
    ElseIf strFormName = "Sub Department" Then
        frmSubDepartment.Show vbModal
    ElseIf strFormName = "Holiday" Then
        frmHoliday.Show vbModal
    ElseIf strFormName = "Shift" Then
        frmShift.Show vbModal
    ElseIf strFormName = "Product Purchase" Then
        frmProduct.Show vbModal
    ElseIf strFormName = "Vehicle" Then
        frmVehicle.Show vbModal
    ElseIf strFormName = "Guesthouse" Then
        frmGuesthouse.Show vbModal
    ElseIf strFormName = "Employee" Then
        frmEmployee.Show vbModal
    ElseIf strFormName = "Guesthouse Booking" Then
        frmGuesthouseBooking.Show vbModal
    ElseIf strFormName = "Issue Bus Pass" Then
        frmIssueBusPass.Show vbModal
    ElseIf strFormName = "Adv. Booking for Vehicle" Then
        frmVehicleBooking.Show vbModal
    ElseIf strFormName = "Employee Shift Sch." Then
        frmEmployeeShiftSchedule.Show vbModal
    ElseIf strFormName = "Vehicle Hire" Then
        frmVehicleHire.Show vbModal
    ElseIf strFormName = "Requirement Entry" Then
        frmRequirement.Show vbModal
    ElseIf strFormName = "Receive Ordered Product" Then
        frmOrderReceive.Show vbModal
    ElseIf strFormName = "Payment of Bills" Then
        frmPayBill.Show vbModal
    ElseIf strFormName = "Change Password" Then
        frmChangePassword.Show vbModal
    ElseIf strFormName = "Vehicle Usage" Then
        frmVehicleUse.Show vbModal
    ElseIf strFormName = "Employee Info" Then
        frmEmployeeInfo.Show vbModal
    ElseIf strFormName = "Employee Leave Saction" Then
        frmLeave.Show vbModal
    ElseIf strFormName = "Item Type" Then
        frmCanteenItemType.Show vbModal
    ElseIf strFormName = "Item" Then
        frmCanteenItem.Show vbModal
    ElseIf strFormName = "New Bill" Then
        frmCanteenBill.Show vbModal
    ElseIf strFormName = "Monthly Shift Schedule" Then
        frmMonthlyShiftSchedule.Show vbModal
    ElseIf strFormName = "Time Card" Then
        frmTimeCard.Show vbModal
    End If
End Sub
