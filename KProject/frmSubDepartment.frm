VERSION 5.00
Begin VB.Form frmSubDepartment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sub Department Master"
   ClientHeight    =   3195
   ClientLeft      =   3045
   ClientTop       =   2580
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Sub Department Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   390
      TabIndex        =   3
      Top             =   360
      Width           =   5265
      Begin VB.CommandButton cmdList 
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
         Left            =   4320
         TabIndex        =   1
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtSubDeptName 
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
         Left            =   1635
         TabIndex        =   2
         Top             =   1350
         Width           =   3435
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H80000000&
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
         Left            =   1635
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   900
         Width           =   2625
      End
      Begin VB.TextBox txtId 
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
         Left            =   1635
         TabIndex        =   4
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   90
         TabIndex        =   8
         Top             =   1410
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Deapartment:"
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
         Left            =   315
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   6
         Top             =   540
         Width           =   240
      End
   End
   Begin KProject.ActionButton ActionButton1 
      Height          =   495
      Left            =   420
      TabIndex        =   0
      Top             =   2490
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
Attribute VB_Name = "frmSubDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim iDeptId As Integer

Private Sub cmdList_Click()
    gBlListDept = True
    frmList.strSql = "select dept_id as ID,dept_name as Department from m_department where status=1"
    frmList.Show vbModal
    txtName.Locked = False
    txtName = gstrListDeptName
    txtName.Locked = True
    iDeptId = gintListDeptId
End Sub

Private Sub ActionButton1_AddClick(iNextNo As Integer)

    Call mClearControl
    txtId = iNextNo + 1
    Call mResetControl(True)
    cmdList.SetFocus
End Sub

Private Sub ActionButton1_CancelClick()
    Call mClearControl
    Call mResetControl(False)
End Sub

Private Sub ActionButton1_DeleteClick()
    With ActionButton1
        .iModifyRecord = txtId
        .strDelete = txtSubDeptName
    End With
    Call mClearControl
End Sub

Private Sub ActionButton1_ModifyClick()
    frmSelect.strSql = "select sub_dept_id as ID ,dept_name as Department, sub_dept_name as SubDepartmentname from m_Sub_department, m_department" & _
                        " where m_sub_department.status=1 and m_sub_department.dept_id=m_Department.dept_id"
    gBlSubDepartment = True
    frmSelect.Show vbModal
    
    If gIntSubDepartment > 0 Then
        Dim rsDisplay As New ADODB.Recordset
        
        rsDisplay.Open "select sub_dept_id,dept_name,sub_dept_name from m_Sub_department,m_department where m_sub_department.dept_id=" & gIntSubDepartment _
                         & " and m_department.dept_id = m_sub_department.dept_id ", gStrConnectionString, adOpenKeyset, adLockOptimistic
        txtId = rsDisplay.Fields(0)
        txtName = rsDisplay.Fields(1)
        txtSubDeptName = rsDisplay.Fields(2)
        iDeptId = gIntSubDepartment
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
        MsgBox "Department Name can not be left blank" & vbCrLf & "Sub department can not exist without Department", vbInformation, "Update"
        Exit Sub
    ElseIf Len(Trim(txtSubDeptName)) = 0 Then
        MsgBox "Sub department Name can not be left blank", vbInformation, "Update"
        Exit Sub
    End If

    ActionButton1.blSave = True
    ActionButton1.SaveSql = "Insert into m_sub_department values (" & txtId & "," & iDeptId & ",'" & txtSubDeptName & "',1)"
    Call mResetControl(False)
End Sub

Private Sub Form_Load()

    Call gFormCenter(Me)
    With ActionButton1
        .ConnectionString = gStrConnectionString
        .PrimaryKeyField = "sub_Dept_id"
        .Table = "m_sub_Department"
    End With
    Call mResetControl(False)
End Sub

Public Sub mResetControl(ByVal blEnable As Boolean)
    cmdList.Enabled = blEnable
    txtSubDeptName.Enabled = blEnable
End Sub

Public Sub mClearControl()
    txtId.Text = ""
    txtName.Text = ""
    txtSubDeptName = ""
End Sub


