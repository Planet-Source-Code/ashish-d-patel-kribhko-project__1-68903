VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLeave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Leave"
   ClientHeight    =   4065
   ClientLeft      =   2115
   ClientTop       =   2955
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Leave Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   180
      TabIndex        =   0
      Top             =   210
      Width           =   6525
      Begin VB.CommandButton cmdAuthorisedBy 
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
         Left            =   5595
         TabIndex        =   16
         Top             =   2250
         Width           =   735
      End
      Begin VB.TextBox txtAuthorisedBy 
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
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2250
         Width           =   3735
      End
      Begin VB.CommandButton cmdEmployeeList 
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
         Left            =   5610
         TabIndex        =   4
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtName 
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
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   900
         Width           =   3765
      End
      Begin VB.TextBox txtId 
         BackColor       =   &H80000018&
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
         Height          =   360
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   450
         Width           =   825
      End
      Begin VB.TextBox txtReason 
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
         Left            =   1770
         TabIndex        =   1
         Top             =   1800
         Width           =   4575
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   360
         Left            =   4830
         TabIndex        =   5
         Top             =   1350
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   19726337
         CurrentDate     =   38821
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   360
         Left            =   1770
         TabIndex        =   6
         Top             =   1350
         Width           =   1635
         _ExtentX        =   2884
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
         Format          =   19726337
         CurrentDate     =   38821
      End
      Begin MSComCtl2.DTPicker dtpRequestDate 
         Height          =   360
         Left            =   4830
         TabIndex        =   7
         Top             =   450
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   19726337
         CurrentDate     =   38821
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Authorised By:"
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
         Left            =   390
         TabIndex        =   17
         Top             =   2310
         Width           =   1290
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "To:"
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
         Left            =   4425
         TabIndex        =   13
         Top             =   1380
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Date:"
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
         Left            =   4260
         TabIndex        =   12
         Top             =   510
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "From:"
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
         Index           =   5
         Left            =   1200
         TabIndex        =   11
         Top             =   1380
         Width           =   510
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Reason:"
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
         Left            =   945
         TabIndex        =   10
         Top             =   1830
         Width           =   765
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Emp Name:"
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
         Left            =   645
         TabIndex        =   9
         Top             =   945
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "No.:"
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
         Left            =   1350
         TabIndex        =   8
         Top             =   510
         Width           =   360
      End
   End
   Begin KProject.ActionButton ActionButton1 
      Height          =   495
      Left            =   840
      TabIndex        =   14
      Top             =   3180
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmLeave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iAuthorisedBy As Integer, iEmpId As Integer

Private Sub cmdAuthorisedBy_Click()
    gBlListEmployee = True
    frmList.strSql = "select emp_id as ID,emp_fname as [First Name],emp_lname as [Last Name],dept_name as Department from view_employee where status=1"
    frmList.Show vbModal
    txtAuthorisedBy.Locked = False
    txtAuthorisedBy = gstrListEmployee
    iAuthorisedBy = gintListEmployee
    txtAuthorisedBy.Locked = True
    gBlListEmployee = False
End Sub

Private Sub cmdEmployeeList_Click()
    gBlListEmployee = True
    frmList.strSql = "select emp_id as ID,emp_fname as [First Name],emp_lname as [Last Name],dept_name as Department from view_employee where status=1"
    frmList.Show vbModal
    txtName.Locked = False
    txtName = gstrListEmployee
    txtName.Locked = True
    iEmpId = gintListEmployee
    gBlListEmployee = False
End Sub


Private Sub ActionButton1_AddClick(iNextNo As Integer)

    Call mClearControl
    txtId = iNextNo + 1
    Call mResetControl(True)
    txtName.SetFocus
End Sub

Private Sub ActionButton1_CancelClick()
    Call mClearControl
    Call mResetControl(False)
End Sub

Private Sub ActionButton1_DeleteClick()
    With ActionButton1
        .iModifyRecord = txtId
        .strDelete = txtName
    End With
    Call mClearControl
End Sub

Private Sub ActionButton1_ModifyClick()
    frmSelect.strSql = "select lr_id as ID ,lr_date as [Leave Date],employee from view_leaveregistration" & _
                        " where status=1"
    gBlLeave = True
    frmSelect.Show vbModal
    
    If gIntLeave > 0 Then
        Dim rsDisplay As New ADODB.Recordset
        
        rsDisplay.Open "select * from view_leaveregistration where lr_id=" & gIntLeave, _
                        gStrConnectionString, adOpenKeyset, adLockOptimistic
        txtId = rsDisplay.Fields(0)
        txtName = rsDisplay.Fields("employee")
        dtpRequestDate.Value = rsDisplay(1)
        dtpFromDate.Value = rsDisplay.Fields(3)
        dtpToDate.Value = rsDisplay.Fields(4)
        txtReason = rsDisplay.Fields(5)
        txtAuthorisedBy = rsDisplay.Fields("Authorised")
        iEmpId = rsDisplay.Fields("emp_id")
        iAuthorisedBy = rsDisplay.Fields("appr_emp_id")
        ActionButton1.blModify = True
        ActionButton1.iModifyRecord = txtId
        Call mResetControl(True)
    Else
        ActionButton1.blModify = False
        ActionButton1.blSave = False
        Call ActionButton1_CancelClick
    End If
End Sub

Private Sub ActionButton1_UpdateClick()
    
    If Len(Trim(txtName)) = 0 Then
        MsgBox "Please select employee name", vbInformation, "Update"
        cmdEmployeeList.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtReason)) = 0 Then
        MsgBox "Reason can not be left blank", vbInformation, "Update"
        txtReason.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtAuthorisedBy)) = 0 Then
        MsgBox "Please select Authorised person name", vbInformation, "Update"
        cmdAuthorisedBy.SetFocus
        Exit Sub
    End If

    ActionButton1.blSave = True
    ActionButton1.SaveSql = "Insert into t_leave_registration values (" & txtId & ",#" & _
        dtpRequestDate.Value & "#," & iEmpId & ",#" & dtpFromDate.Value & "#,#" & dtpToDate.Value & _
        "#,'" & txtReason & "'," & iAuthorisedBy & ",1)"
    Call mResetControl(False)
End Sub

Private Sub Form_Load()

    Call gFormCenter(Me)
    With ActionButton1
        .ConnectionString = gStrConnectionString
        .PrimaryKeyField = "lr_id"
        .Table = "t_leave_registration"
    End With
    Call mResetControl(False)
End Sub

Public Sub mResetControl(ByVal blEnable As Boolean)
    txtId.Enabled = blEnable
    txtName.Enabled = blEnable
    dtpRequestDate.Enabled = blEnable
    dtpFromDate.Enabled = blEnable
    dtpToDate.Enabled = blEnable
    txtReason.Enabled = blEnable
    txtAuthorisedBy.Enabled = blEnable
    cmdEmployeeList.Enabled = blEnable
    cmdAuthorisedBy.Enabled = blEnable
End Sub

Public Sub mClearControl()
    txtId.Locked = False
    txtName.Locked = False
    txtAuthorisedBy.Locked = False
    txtId.Text = ""
    txtName.Text = ""
    dtpFromDate.Value = Date
    dtpToDate.Value = Date
    txtReason.Text = ""
    txtAuthorisedBy.Text = ""
    dtpRequestDate.Value = Date
    txtId.Locked = True
    txtName.Locked = True
    txtAuthorisedBy.Locked = True
End Sub


