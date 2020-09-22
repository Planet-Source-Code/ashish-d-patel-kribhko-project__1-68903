VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmIssueBusPass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Issue Bus Pass ( Free and Chargable )"
   ClientHeight    =   4005
   ClientLeft      =   2055
   ClientTop       =   2580
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Issue Detail"
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
      TabIndex        =   9
      Top             =   270
      Width           =   6525
      Begin VB.TextBox txtSerialNo 
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
         TabIndex        =   3
         Top             =   1350
         Width           =   1605
      End
      Begin VB.ComboBox cmbPassType 
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
         ItemData        =   "frmIssueBusPass.frx":0000
         Left            =   4830
         List            =   "frmIssueBusPass.frx":000A
         TabIndex        =   4
         Top             =   1350
         Width           =   1515
      End
      Begin VB.TextBox txtDependent 
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
         TabIndex        =   5
         Top             =   1770
         Width           =   1605
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
         Left            =   1770
         TabIndex        =   11
         Top             =   450
         Width           =   825
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
         TabIndex        =   10
         Top             =   900
         Width           =   3765
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
         TabIndex        =   2
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtRelation 
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
         Left            =   4830
         TabIndex        =   6
         Top             =   1785
         Width           =   1515
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   360
         Left            =   4830
         TabIndex        =   8
         Top             =   2220
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
         TabIndex        =   7
         Top             =   2220
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
         TabIndex        =   1
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pass Type:"
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
         Index           =   3
         Left            =   3720
         TabIndex        =   20
         Top             =   1380
         Width           =   1035
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
         Left            =   1335
         TabIndex        =   19
         Top             =   510
         Width           =   360
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
         Left            =   615
         TabIndex        =   18
         Top             =   945
         Width           =   1065
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dependent Name:"
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
         Left            =   30
         TabIndex        =   17
         Top             =   1800
         Width           =   1650
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valid From:"
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
         Left            =   630
         TabIndex        =   16
         Top             =   2280
         Width           =   1020
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
         TabIndex        =   15
         Top             =   510
         Width           =   480
      End
      Begin VB.Label Label4 
         Caption         =   "Relation:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   14
         Top             =   1800
         Width           =   825
      End
      Begin VB.Label Label5 
         Caption         =   "Valid To:"
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
         Left            =   3930
         TabIndex        =   13
         Top             =   2250
         Width           =   945
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Serial No.:"
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
         Left            =   750
         TabIndex        =   12
         Top             =   1395
         Width           =   930
      End
   End
   Begin KProject.ActionButton ActionButton1 
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   3240
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
Attribute VB_Name = "frmIssueBusPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub ActionButton1_AddClick(iNextNo As Integer)

    Call mClearControl
    txtId = iNextNo + 1
    Call mResetControl(True)
    dtpRequestDate.SetFocus
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
    frmSelect.strSql = "select ibp_id as ID ,ibp_date as [Issue Date],emp_fname as [To Employee] from view_issuebuspass" & _
                        " where status=1"
    gBlIssuePass = True
    frmSelect.Show vbModal
    
    If gIntIssuePass > 0 Then
        Dim rsDisplay As New ADODB.Recordset
        
        rsDisplay.Open "select * from view_issuebuspass where ibp_id=" & gIntIssuePass, _
                        gStrConnectionString, adOpenKeyset, adLockOptimistic
        txtId = rsDisplay.Fields(0)
        dtpRequestDate.Value = rsDisplay.Fields(1)
        txtSerialNo = rsDisplay(2)
        txtName = rsDisplay.Fields(3)
        cmbPassType.Text = rsDisplay.Fields(4)
        txtDependent = rsDisplay.Fields(5)
        txtRelation = rsDisplay.Fields(6)
        dtpFromDate.Value = rsDisplay.Fields(7)
        dtpToDate.Value = rsDisplay.Fields(8)
        gintListEmployee = rsDisplay.Fields("emp_id")
        
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
        MsgBox "Please select Employee Name", vbInformation, "Update"
        cmdEmployeeList.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtSerialNo)) = 0 Then
        MsgBox "Please enter Serial No. of Pass ", vbInformation, "Update"
        txtSerialNo.SetFocus
        Exit Sub
    ElseIf Len(Trim(cmbPassType.Text)) = 0 Then
        MsgBox "Please select Pass Type", vbInformation, "Update"
        cmbPassType.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtDependent)) = 0 Then
        MsgBox "Please enter Dependent Name", vbInformation, "Update"
        txtDependent.SetFocus
        Exit Sub
    End If

    ActionButton1.blSave = True
    'need to replace zero with guserid if guserid is admin
    ActionButton1.SaveSql = "Insert into t_issue_bus_pass values (" & txtId & _
            ",#" & dtpRequestDate & "#,'" & txtSerialNo & "'," & gintListEmployee & _
            ",'" & cmbPassType.Text & "','" & txtDependent & "','" & txtRelation & _
            "',#" & dtpFromDate.Value & "#,#" & dtpToDate.Value & "#,1)"
    Call mResetControl(False)
End Sub

Private Sub cmbPassType_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdEmployeeList_Click()
    gBlListEmployee = True
    frmList.strSql = "select emp_id as ID,emp_fname as [First Name],emp_lname as [Last Name],dept_name as Department from view_employee where status=1"
    frmList.Show vbModal
    txtName.Locked = False
    txtName = gstrListEmployee
    txtName.Locked = True
    gBlListEmployee = False
End Sub

Private Sub Form_Load()

    Call gFormCenter(Me)
    With ActionButton1
        .ConnectionString = gStrConnectionString
        .PrimaryKeyField = "ibp_id"
        .Table = "t_issue_bus_pass"
    End With
    Call mResetControl(False)
End Sub

Public Sub mResetControl(ByVal blEnable As Boolean)
    txtName.Enabled = blEnable
    dtpRequestDate.Enabled = blEnable
    txtSerialNo.Enabled = blEnable
    txtDependent.Enabled = blEnable
    cmbPassType.Enabled = blEnable
    dtpFromDate.Enabled = blEnable
    dtpToDate.Enabled = blEnable
    cmdEmployeeList.Enabled = blEnable
    txtRelation.Enabled = blEnable
End Sub

Public Sub mClearControl()
    txtId.Text = ""
    txtName.Locked = False
    txtName.Text = ""
    txtName.Locked = True
    dtpRequestDate.Value = Date
    txtSerialNo.Text = ""
    txtDependent.Text = ""
    txtRelation.Text = ""
    cmbPassType.Text = ""
    dtpFromDate.Value = Date
    dtpToDate.Value = Date
End Sub

