VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEmployeeShiftSchedule 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employee Shift Schedule"
   ClientHeight    =   6210
   ClientLeft      =   2235
   ClientTop       =   1470
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3120
      TabIndex        =   8
      Top             =   5640
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Caption         =   "Shift Assign"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5355
      Left            =   150
      TabIndex        =   10
      Top             =   180
      Width           =   6705
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   360
         Left            =   4950
         TabIndex        =   23
         Top             =   4020
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   19791873
         CurrentDate     =   38830
      End
      Begin VB.TextBox txtRemark 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1650
         TabIndex        =   7
         Top             =   4830
         Width           =   4845
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "&Select All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5220
         TabIndex        =   3
         Top             =   1200
         Width           =   1305
      End
      Begin VB.CommandButton cmdShiftList 
         Caption         =   "List..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3270
         TabIndex        =   6
         Top             =   4410
         Width           =   735
      End
      Begin VB.CommandButton cmdSubDepartmentList 
         Caption         =   "List..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3660
         TabIndex        =   2
         Top             =   1170
         Width           =   735
      End
      Begin VB.CommandButton cmdDepartmentList 
         Caption         =   "List..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3660
         TabIndex        =   1
         Top             =   750
         Width           =   735
      End
      Begin VB.TextBox txtShiftname 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   4410
         Width           =   1545
      End
      Begin VB.ListBox lstEmployee 
         Height          =   2310
         ItemData        =   "frmEmployeeShiftSchedule.frx":0000
         Left            =   1650
         List            =   "frmEmployeeShiftSchedule.frx":0002
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   1620
         Width           =   4875
      End
      Begin VB.TextBox txtSubDepartment 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1170
         Width           =   1935
      End
      Begin VB.TextBox txtDepartment 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   750
         Width           =   1935
      End
      Begin VB.TextBox txtId 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   330
         Width           =   1245
      End
      Begin MSComCtl2.DTPicker dtpEffectiveDate 
         Height          =   360
         Left            =   1650
         TabIndex        =   5
         Top             =   3990
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   19791873
         CurrentDate     =   38822
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "To Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   4050
         TabIndex        =   22
         Top             =   4080
         Width           =   780
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Remark:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   780
         TabIndex        =   21
         Top             =   4920
         Width           =   765
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Shift Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   525
         TabIndex        =   16
         Top             =   4470
         Width           =   1020
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "From Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   555
         TabIndex        =   15
         Top             =   4050
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Employees:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   465
         TabIndex        =   14
         Top             =   1680
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "SubDepartment:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   13
         Top             =   1230
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Department:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   450
         TabIndex        =   12
         Top             =   810
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1305
         TabIndex        =   11
         Top             =   360
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Canc&el"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4170
      TabIndex        =   9
      Top             =   5640
      Width           =   1005
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2070
      TabIndex        =   0
      Top             =   5640
      Width           =   1005
   End
End
Attribute VB_Name = "frmEmployeeShiftSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Public Sub mClearControl()
    
    txtId.Locked = False
    txtId.Text = ""
    txtId.Locked = True
    txtDepartment.Locked = False
    txtDepartment.Text = ""
    txtDepartment.Locked = True
    txtSubDepartment.Locked = False
    txtSubDepartment.Text = ""
    txtSubDepartment.Locked = True
    lstEmployee.Clear
    dtpEffectiveDate.Value = Date
    dtpToDate.Value = Date
    txtShiftname.Locked = False
    txtShiftname.Text = ""
    txtShiftname.Locked = True
    cmdSelectAll.Caption = "&Select All"
    txtRemark.Text = ""
    
End Sub

Public Sub mResetControl(blEnable As Boolean)
    txtId.Enabled = blEnable
    txtDepartment.Enabled = blEnable
    txtSubDepartment.Enabled = blEnable
    lstEmployee.Enabled = blEnable
    dtpEffectiveDate.Enabled = blEnable
    dtpToDate.Enabled = blEnable
    txtShiftname.Enabled = blEnable
    txtRemark.Enabled = blEnable
    cmdDepartmentList.Enabled = blEnable
    cmdSubDepartmentList.Enabled = blEnable
    cmdShiftList.Enabled = blEnable
    cmdSelectAll.Enabled = blEnable
End Sub

Private Sub cmdAdd_Click()
On Error GoTo AddError
Dim rsAdd As New ADODB.Recordset

    cmdUpdate.Enabled = True
    cmdAdd.Enabled = False
    Call mClearControl
    Call mResetControl(True)
    
    rsAdd.Open "select max(ess_id) from t_employee_shift_schedule", gStrConnectionString, adOpenKeyset, adLockOptimistic
    
    If rsAdd.RecordCount > 0 Then
        txtId.Locked = False
        txtId.Text = rsAdd.Fields(0) + 1
        txtId.Locked = True
    End If
    Exit Sub
AddError:
    MsgBox "Error: " & Err.Number & vbCrLf & "Desc :" & Err.Description, vbCritical, "Add Error"
    Err.Clear
End Sub

Private Sub cmdCancel_Click()

    Call mClearControl
    Call mResetControl(False)
    cmdAdd.Enabled = True
    cmdUpdate.Enabled = False
End Sub

Private Sub cmdDepartmentList_Click()
    gBlListDept = True
    frmList.strSql = "select dept_id as ID,dept_name as Department from m_department where status=1"
    frmList.Show vbModal
    txtDepartment.Locked = False
    txtDepartment = gstrListDeptName
    txtDepartment.Locked = True
    gBlListDept = False

End Sub

Private Sub cmdSelectAll_Click()
    Dim i As Integer, blAction As Boolean
    
    If cmdSelectAll.Caption = "&Select All" Then
        blAction = True
    Else
        blAction = False
    End If
    
    For i = 0 To lstEmployee.ListCount - 1
        lstEmployee.Selected(i) = blAction
    Next
    
    If blAction Then
        cmdSelectAll.Caption = "&Deselect All"
    Else
        cmdSelectAll.Caption = "&Select All"
    End If
End Sub

Private Sub cmdShiftList_Click()
    gBlListShift = True
    frmList.strSql = "select shf_id as ID,shf_name as Shift,shf_start_time as Start from m_shift where status=1"
    frmList.Show vbModal
    txtShiftname.Locked = False
    txtShiftname = gstrListShiftName
    txtShiftname.Locked = True
    gBlListShift = False

End Sub

Private Sub cmdSubDepartmentList_Click()
    gBlListSubDept = True
    frmList.strSql = "select sub_dept_id as ID,dept_name as Department,sub_dept_name as SubDepartment " & _
                    " from m_department md, m_sub_department ms where ms.status=1" & _
                    " and md.dept_id=ms.dept_id and ms.dept_id=" & gintListDeptId
    frmList.Show vbModal
    txtSubDepartment.Locked = False
    txtSubDepartment = gstrListSubDeptName
    txtSubDepartment.Locked = True
    gBlListSubDept = False
    
    Call mFillList(gintListDeptId, gintListSubDeptId)
End Sub


Private Sub cmdUpdate_Click()
On Error GoTo UpdateError
    Dim cnUpdate As New ADODB.Connection, strSql As String, i As Integer, iSelect As Integer
    Dim iCnt As Integer
    
    If Len(Trim(txtDepartment)) = 0 Then
        MsgBox "Please select Department", vbInformation, "Update"
        cmdDepartmentList.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtShiftname)) = 0 Then
        MsgBox "Please select New Shift", vbInformation, "Update"
        cmdShiftList.SetFocus
        Exit Sub
    Else
        For i = 0 To lstEmployee.ListCount - 1
            If lstEmployee.Selected(i) = True Then
                iSelect = iSelect + 1
            End If
        Next
        If iSelect = 0 Then
            MsgBox "No employee selected to change shift", vbInformation, "Update"
            lstEmployee.SetFocus
            Exit Sub
        End If
    End If

    cnUpdate.ConnectionString = gStrConnectionString
    cnUpdate.Open
    
    cnUpdate.BeginTrans
    iCnt = txtId
        For i = 0 To lstEmployee.ListCount - 1
            If lstEmployee.Selected(i) = True Then
                strSql = "Insert into t_Employee_shift_schedule values (" & iCnt & _
                        "," & lstEmployee.ItemData(i) & "," & gintListShiftId & ",#" & _
                        dtpEffectiveDate.Value & "#,#" & dtpToDate.Value & "#,'" & txtRemark & "',1)"
                cnUpdate.Execute strSql
'                strSql = "Update m_employee set shf_id=" & gintListShiftId & _
'                        " where emp_id = " & lstEmployee.ItemData(i)
'                cnUpdate.Execute strSql
            End If
            iCnt = iCnt + 1
        Next
    cnUpdate.CommitTrans
    MsgBox "Record(s) Updated ", vbInformation, "Update"
    cmdCancel_Click
Exit Sub
UpdateError:
    MsgBox "Error: " & Err.Number & vbCrLf & "Desc :" & Err.Description, vbCritical, "Update Error"
    Err.Clear
End Sub

Private Sub Form_Load()
    Call gFormCenter(Me)
    Call mClearControl
    Call mResetControl(False)
End Sub

Public Sub mFillList(iDepartment As Integer, iSubDepartment As Integer)
On Error GoTo FillError
    Dim rsFill As New ADODB.Recordset, i As Integer
    
    rsFill.Open "select Emp_id,Emp_fname +' ' + emp_mname + ' ' + emp_lname from m_employee " & _
                " where dept_id=" & gintListDeptId & " and sub_dept_id=" & gintListSubDeptId & _
                " and status=1", gStrConnectionString, adOpenKeyset, adLockOptimistic
    
    If rsFill.RecordCount > 0 Then
        lstEmployee.Clear
        For i = 0 To rsFill.RecordCount - 1
            lstEmployee.AddItem rsFill.Fields(1)
            lstEmployee.ItemData(i) = rsFill.Fields(0)
            rsFill.MoveNext
        Next
    End If
Exit Sub
FillError:
    MsgBox "Error: " & Err.Number & vbCrLf & "Desc :" & Err.Description, vbCritical, "Update Error"
    Err.Clear
End Sub
