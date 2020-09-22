VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPayBill 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pay Bill"
   ClientHeight    =   5250
   ClientLeft      =   2085
   ClientTop       =   1725
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Bill Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4155
      Left            =   210
      TabIndex        =   10
      Top             =   240
      Width           =   5955
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
         Left            =   1680
         TabIndex        =   9
         Top             =   3630
         Width           =   4095
      End
      Begin VB.TextBox txtID 
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   270
         Width           =   1245
      End
      Begin VB.TextBox txtEmployee 
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   690
         Width           =   3285
      End
      Begin VB.TextBox txtBillNo 
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
         Left            =   1680
         TabIndex        =   2
         Top             =   1110
         Width           =   1245
      End
      Begin VB.ComboBox cmbBillType 
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
         ItemData        =   "frmPayBill.frx":0000
         Left            =   1680
         List            =   "frmPayBill.frx":0010
         TabIndex        =   4
         Top             =   1530
         Width           =   1275
      End
      Begin VB.TextBox txtCompanyName 
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
         Left            =   1680
         TabIndex        =   5
         Top             =   1950
         Width           =   4095
      End
      Begin VB.TextBox txtOwnerName 
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
         Left            =   1680
         TabIndex        =   6
         Top             =   2370
         Width           =   4095
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
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
         Left            =   1680
         TabIndex        =   7
         Top             =   2790
         Width           =   1245
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   3210
         Width           =   3255
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
         Left            =   5070
         TabIndex        =   1
         Top             =   690
         Width           =   735
      End
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
         Left            =   5040
         TabIndex        =   8
         Top             =   3210
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtpBillDate 
         Height          =   360
         Left            =   4290
         TabIndex        =   3
         Top             =   1110
         Width           =   1515
         _ExtentX        =   2672
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
         CurrentDate     =   38826
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   1
         Left            =   840
         TabIndex        =   23
         Top             =   3660
         Width           =   765
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
         Left            =   1380
         TabIndex        =   22
         Top             =   330
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Employee:"
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
         Left            =   630
         TabIndex        =   21
         Top             =   750
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "BillDate:"
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
         Left            =   3480
         TabIndex        =   20
         Top             =   1170
         Width           =   750
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bill No.:"
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
         Left            =   930
         TabIndex        =   19
         Top             =   1170
         Width           =   675
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Company Name:"
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
         TabIndex        =   18
         Top             =   2010
         Width           =   1515
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Owner Name:"
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
         Top             =   2430
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Amount:"
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
         Left            =   885
         TabIndex        =   16
         Top             =   2850
         Width           =   720
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
         Left            =   315
         TabIndex        =   15
         Top             =   3270
         Width           =   1290
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bill Type:"
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
         Left            =   765
         TabIndex        =   14
         Top             =   1590
         Width           =   840
      End
   End
   Begin KProject.ActionButton ActionButton1 
      Height          =   495
      Left            =   630
      TabIndex        =   0
      Top             =   4590
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
Attribute VB_Name = "frmPayBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim iEmployeeid As Integer, iAuthorisedId As Integer

Private Sub ActionButton1_AddClick(iNextNo As Integer)

    Call mClearControl
    txtId.Locked = False
    txtId = iNextNo + 1
    txtId.Locked = True
    Call mResetControl(True)
    cmdEmployeeList.SetFocus
End Sub

Private Sub ActionButton1_CancelClick()
    Call mClearControl
    Call mResetControl(False)
End Sub

Private Sub ActionButton1_DeleteClick()
    With ActionButton1
        .iModifyRecord = txtId
        .strDelete = txtBillNo
    End With
    Call mClearControl
End Sub

Private Sub ActionButton1_ModifyClick()
    frmSelect.strSql = "select pb_id as ID ,pb_billno as BillNum,pb_date as [Bill Date],emp_fname as Employee from view_PaymentOfBills" & _
                        " where status=1"
    gBlPaymentOfBill = True
    frmSelect.Show vbModal
    
    If gIntPaymentOfBill > 0 Then
        Dim rsDisplay As New ADODB.Recordset
        
        rsDisplay.Open "select * from view_Paymentofbills where pb_id=" & gIntPaymentOfBill, _
                        gStrConnectionString, adOpenKeyset, adLockOptimistic
        txtId.Locked = False
        txtEmployee.Locked = False
        txtAuthorisedBy.Locked = False
        txtId = rsDisplay.Fields(0)
        dtpBillDate.Value = rsDisplay.Fields(1)
        txtEmployee.Text = rsDisplay.Fields(2)
        txtBillNo = rsDisplay.Fields(3)
        cmbBillType.Text = rsDisplay.Fields(4)
        txtCompanyName.Text = rsDisplay.Fields(5)
        txtOwnerName.Text = rsDisplay.Fields(6)
        txtRemark.Text = rsDisplay.Fields(7)
        txtAmount.Text = rsDisplay.Fields(8)
        txtAuthorisedBy.Text = rsDisplay.Fields(9)
        txtId.Locked = True
        txtEmployee.Locked = True
        txtAuthorisedBy.Locked = True
        iEmployeeid = rsDisplay.Fields("emp_id")
        iAuthorisedId = rsDisplay.Fields("auth_emp_id")
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
    
    If Len(Trim(txtEmployee)) = 0 Then
        MsgBox "Please select employee", vbInformation, "Update"
        cmdEmployeeList.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtBillNo)) = 0 Then
        MsgBox "Bill no. is required", vbInformation, "Update"
        txtBillNo.SetFocus
        Exit Sub
    ElseIf Len(Trim(cmbBillType.Text)) = 0 Then
        MsgBox "Bill Type can not be left blank", vbInformation, "Update"
        cmbBillType.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtCompanyName.Text)) = 0 Then
        MsgBox "Please enter company name", vbInformation, "Update"
        txtCompanyName.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtOwnerName.Text)) = 0 Then
        MsgBox "Owner name is required", vbInformation, "Update"
        txtOwnerName.SetFocus
        Exit Sub
    ElseIf Val(txtAmount.Text) = 0 Then
        MsgBox "Amount must be greater than 0", vbInformation, "Update"
        txtAmount.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtAuthorisedBy.Text)) = 0 Then
        MsgBox "Please select Authorised person name", vbInformation, "Update"
        cmdAuthorisedBy.SetFocus
        Exit Sub
    End If

    ActionButton1.blSave = True
    ActionButton1.SaveSql = "Insert into t_payment_of_bills values (" & txtId & ",#" & _
            dtpBillDate.Value & "#," & iEmployeeid & ",'" & txtBillNo.Text & "','" & _
            cmbBillType.Text & "','" & txtCompanyName.Text & "','" & txtOwnerName & _
            "','" & txtRemark & "'," & CCur(txtAmount) & "," & iAuthorisedId & ",1)"
    Call mResetControl(False)
End Sub

Private Sub cmdAuthorisedBy_Click()
    gBlListEmployee = True
    frmList.strSql = "select emp_id as ID,emp_fname as [First Name],emp_lname as [Last Name],dept_name as Department from view_employee where status=1"
    frmList.Show vbModal
    txtAuthorisedBy.Locked = False
    txtAuthorisedBy = gstrListEmployee
    iAuthorisedId = gintListEmployee
    txtAuthorisedBy.Locked = True
    gBlListEmployee = False
End Sub

Private Sub cmdEmployeeList_Click()
    gBlListEmployee = True
    frmList.strSql = "select emp_id as ID,emp_fname as [First Name],emp_lname as [Last Name],dept_name as Department from view_employee where status=1"
    frmList.Show vbModal
    txtEmployee.Locked = False
    txtEmployee = gstrListEmployee
    iEmployeeid = gintListEmployee
    txtEmployee.Locked = True
    gBlListEmployee = False
End Sub

Private Sub Form_Load()

    Call gFormCenter(Me)
    dtpBillDate.Value = Date
    With ActionButton1
        .ConnectionString = gStrConnectionString
        .PrimaryKeyField = "pb_id"
        .Table = "t_payment_of_bills"
    End With
    Call mResetControl(False)
End Sub

Public Sub mResetControl(ByVal blEnable As Boolean)
    txtId.Enabled = blEnable
    dtpBillDate.Enabled = blEnable
    txtEmployee.Enabled = blEnable
    txtBillNo.Enabled = blEnable
    cmbBillType.Enabled = blEnable
    txtCompanyName.Enabled = blEnable
    txtRemark.Enabled = blEnable
    txtAmount.Enabled = blEnable
    txtAuthorisedBy.Enabled = blEnable
    cmdEmployeeList.Enabled = blEnable
    cmdAuthorisedBy.Enabled = blEnable
    txtOwnerName.Enabled = blEnable
End Sub

Public Sub mClearControl()
    txtId.Locked = False
    txtEmployee.Locked = False
    txtAuthorisedBy.Locked = False
    txtId.Text = ""
    txtEmployee.Text = ""
    txtBillNo.Text = ""
    dtpBillDate.Value = Date
    cmbBillType.Text = ""
    txtCompanyName.Text = ""
    txtRemark.Text = ""
    txtAmount.Text = ""
    txtAuthorisedBy.Text = ""
    txtId.Locked = True
    txtEmployee.Locked = True
    txtAuthorisedBy.Locked = True
    txtOwnerName.Text = ""
End Sub




