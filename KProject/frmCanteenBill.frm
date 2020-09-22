VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCanteenBill 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Canteen bill"
   ClientHeight    =   7380
   ClientLeft      =   2265
   ClientTop       =   765
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   6930
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Left            =   1410
      TabIndex        =   22
      Top             =   6780
      Width           =   1005
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
      Left            =   3510
      TabIndex        =   21
      Top             =   6780
      Width           =   1005
   End
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
      Left            =   2460
      TabIndex        =   20
      Top             =   6780
      Width           =   1005
   End
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
      Height          =   6495
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   5835
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
         Left            =   4950
         TabIndex        =   29
         Top             =   1590
         Width           =   735
      End
      Begin VB.TextBox txtCoupenNo 
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
         Left            =   1380
         TabIndex        =   28
         Top             =   1170
         Width           =   1245
      End
      Begin VB.ComboBox cmbCategory 
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
         ItemData        =   "frmCanteenBill.frx":0000
         Left            =   1380
         List            =   "frmCanteenBill.frx":000A
         TabIndex        =   26
         Top             =   750
         Width           =   1245
      End
      Begin MSComCtl2.DTPicker dtpBillDate 
         Height          =   360
         Left            =   4020
         TabIndex        =   24
         Top             =   300
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
         Format          =   53805057
         CurrentDate     =   38832
      End
      Begin VB.CommandButton cmdApprovedBy 
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
         Left            =   4950
         TabIndex        =   13
         Top             =   6000
         Width           =   735
      End
      Begin VB.TextBox txtApprovedBy 
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
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   6000
         Width           =   3495
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
         Left            =   4950
         TabIndex        =   11
         Top             =   2010
         Width           =   735
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
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2010
         Width           =   3495
      End
      Begin VB.TextBox txtEmployeeName 
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
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1590
         Width           =   3495
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
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   300
         Width           =   1245
      End
      Begin VB.Frame Frame2 
         Caption         =   "Details "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   90
         TabIndex        =   1
         Top             =   2460
         Width           =   5655
         Begin VB.TextBox txtNo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   210
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   450
            Width           =   615
         End
         Begin VB.ComboBox cmbItemList 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   990
            TabIndex        =   5
            Top             =   450
            Width           =   2505
         End
         Begin VB.TextBox txtQty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3540
            TabIndex        =   4
            Top             =   450
            Width           =   735
         End
         Begin VB.TextBox txtRate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   450
            Width           =   825
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   450
            Width           =   675
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   2595
            Left            =   120
            TabIndex        =   7
            Top             =   810
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   4577
            _Version        =   393216
            Cols            =   6
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   360
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
            _Band(0).ColHeader=   1
         End
      End
      Begin MSComCtl2.DTPicker dtpCoupenDate 
         Height          =   360
         Left            =   4020
         TabIndex        =   14
         Top             =   1170
         Width           =   1605
         _ExtentX        =   2831
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
         Format          =   53805057
         CurrentDate     =   38823
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Coupen No.:"
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
         Index           =   2
         Left            =   210
         TabIndex        =   27
         Top             =   1230
         Width           =   1110
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Category:"
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
         Left            =   435
         TabIndex        =   25
         Top             =   810
         Width           =   870
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Index           =   1
         Left            =   3450
         TabIndex        =   23
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Approved By:"
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
         TabIndex        =   19
         Top             =   6030
         Width           =   1230
      End
      Begin VB.Label Label4 
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
         Left            =   210
         TabIndex        =   18
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Coupen Date:"
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
         Left            =   2700
         TabIndex        =   17
         Top             =   1230
         Width           =   1230
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
         Left            =   330
         TabIndex        =   16
         Top             =   1620
         Width           =   975
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
         Left            =   930
         TabIndex        =   15
         Top             =   390
         Width           =   360
      End
   End
   Begin VB.Menu mnuRight 
      Caption         =   "Right"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove"
      End
   End
End
Attribute VB_Name = "frmCanteenBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim iRow As Integer, blGridClick As Boolean, blMenuAdd As Boolean
Dim iStartTop As Integer
Dim iApprovedBy As Integer
Private Type ItemList
    iItemCode As Integer
    strDescription As String
    cRate As Currency
End Type
Dim tpItemLst() As ItemList

Private Sub cmbCategory_Click()
    If cmbCategory.Text = "GENERAL" Then
        txtDepartment.Enabled = False
        cmdDepartmentList.Enabled = False
        txtEmployeeName.Enabled = True
        cmdEmployeeList.Enabled = True
    ElseIf cmbCategory.Text = "ATR" Then
        txtEmployeeName.Enabled = False
        cmdEmployeeList.Enabled = False
        txtDepartment.Enabled = True
        cmdDepartmentList.Enabled = True
    End If
End Sub

Private Sub cmbCategory_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub

Private Sub cmbCategory_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmbItemList_Click()

    If Len(cmbItemList.Text) > 0 Then
        txtRate.Text = tpItemLst(cmbItemList.ListIndex).cRate
        MSHFlexGrid1.TextMatrix(txtNo.Text, 5) = tpItemLst(cmbItemList.ListIndex).iItemCode
    End If
End Sub

Private Sub cmbItemList_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub cmbItemList_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdAdd_Click()
On Error GoTo AddError
Dim rsAdd As New ADODB.Recordset

    cmdUpdate.Enabled = True
    cmdAdd.Enabled = False
    Call mClearControl
    Call mResetControl(True)
    
    rsAdd.Open "select max(b_no) from t_canteen_bill", gStrConnectionString, adOpenKeyset, adLockOptimistic
    
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

Private Sub cmdApprovedBy_Click()
    gBlListApproveBy = True
    frmList.strSql = "select emp_id as ID,emp_fname as [First Name],emp_lname as [Last Name],dept_name as Department from view_employee where status=1"
    frmList.Show vbModal
    txtApprovedBy.Locked = False
    txtApprovedBy = gstrListApproveBy
    iApprovedBy = gintListApproveBy
    txtApprovedBy.Locked = True
    gBlListApproveBy = False
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
End Sub

Private Sub cmdEmployeeList_Click()
    gBlListEmployee = True
    frmList.strSql = "select emp_id as ID,emp_fname as [First Name],emp_lname as [Last Name],dept_name as Department from view_employee where status=1"
    frmList.Show vbModal
    txtEmployeeName.Locked = False
    txtEmployeeName = gstrListEmployee
    txtEmployeeName.Locked = True
    gBlListEmployee = False
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo UpdateError
    
    If Len(Trim(txtCoupenNo)) = 0 Then
        MsgBox "Coupen No. can not be left blank", vbInformation, "Update"
        txtCoupenNo.SetFocus
        Exit Sub
    ElseIf Len(cmbCategory.Text) = 0 Then
        MsgBox "Please select category", vbInformation, "Update"
        cmbCategory.SetFocus
        Exit Sub
    ElseIf Trim(cmbCategory.Text) = "GENERAL" And Len(Trim(txtEmployeeName)) = 0 Then
        MsgBox "Please select employee", vbInformation, "Update"
        cmdEmployeeList.SetFocus
        Exit Sub
    ElseIf Trim(cmbCategory.Text) = "ATR" And Len(Trim(txtDepartment)) = 0 Then
        MsgBox "Please select Department", vbInformation, "Update"
        cmdDepartmentList.SetFocus
        Exit Sub
    End If
    If Len(txtApprovedBy) = 0 Then
        MsgBox "Please select Approved by", vbInformation, "Update"
        cmdApprovedBy.SetFocus
        Exit Sub
    End If
    
    Dim rsCheck As New ADODB.Recordset, strCheck As String
    Dim cnUpdate As New ADODB.Connection, i As Integer
    
    strCheck = "select * from t_canteen_bill where b_coupen_no='" & txtCoupenNo & "'"
    rsCheck.Open strCheck, gStrConnectionString, adOpenKeyset, adLockOptimistic
    
    If rsCheck.RecordCount > 0 Then
        MsgBox "Coupen already used", vbExclamation, "Update"
        Exit Sub
    End If
    
    cnUpdate.ConnectionString = gStrConnectionString
    cnUpdate.Open
    cnUpdate.BeginTrans
    strCheck = "Insert into t_canteen_bill values (" & txtId.Text & ",#" & dtpBillDate.Value & _
        "#,'" & cmbCategory.Text & "','" & txtCoupenNo.Text & "',#" & dtpCoupenDate.Value & _
        "#," & gintListEmployee & "," & gintListDeptId & "," & iApprovedBy & ",1)"
    cnUpdate.Execute strCheck
    
    With MSHFlexGrid1
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 2)) <> 0 Then
                strCheck = "Insert into t_canteen_bill_detail values (" & txtId.Text & _
                "," & .TextMatrix(i, 0) & "," & .TextMatrix(i, 5) & "," & .TextMatrix(i, 2) & _
                 "," & .TextMatrix(i, 3) & ")"
                cnUpdate.Execute strCheck
            End If
        Next
    End With
    cnUpdate.CommitTrans
    MsgBox "Bill Details Updated successfully", vbInformation, "Update"
    
    With CrystalReport1
        .DataFiles(0) = App.Path & "\KSystemDb.mdb"
        If cmbCategory.Text <> "ATR" Then
            .ReportFileName = App.Path & "\Report\rptEmployeeCanteenBill.rpt"
        Else
            .ReportFileName = App.Path & "\Report\rptDepartmentCanteenBill.rpt"
        End If
        .PrintReport
    End With
    cmdCancel_Click
Exit Sub
UpdateError:
    MsgBox "Error :" & Err.Number & vbCrLf & "Desc :" & Err.Description, vbCritical, "Update Error"
    Err.Clear
End Sub

Private Sub Form_Load()
    
    Call gFormCenter(Me)
    Call mGetItemDetails
    Call mResetControl(False)
    Call mInitGrid
    iStartTop = txtNo.Top
End Sub

Public Sub mResetControl(ByVal blEnable As Boolean)
    txtId.Enabled = blEnable
    cmbCategory.Enabled = blEnable
    txtCoupenNo.Enabled = blEnable
    dtpCoupenDate.Enabled = blEnable
    dtpBillDate.Enabled = blEnable
    txtDepartment.Enabled = blEnable
    txtApprovedBy.Enabled = blEnable
    txtEmployeeName.Enabled = blEnable
    cmdDepartmentList.Enabled = blEnable
    cmdApprovedBy.Enabled = blEnable
    cmdEmployeeList.Enabled = blEnable
    MSHFlexGrid1.Enabled = blEnable
    txtNo.Enabled = blEnable
    cmbItemList.Enabled = blEnable
    txtQty.Enabled = blEnable
    txtRate.Enabled = blEnable
    txtTotal.Enabled = blEnable
End Sub

Public Sub mClearControl()
    txtId.Locked = False
    txtDepartment.Locked = False
    txtApprovedBy.Locked = False
    txtEmployeeName.Locked = False
    txtId.Text = ""
    cmbCategory.Text = ""
    txtDepartment.Text = ""
    txtApprovedBy.Text = ""
    txtEmployeeName.Text = ""
    txtCoupenNo.Text = ""
    txtId.Locked = True
    txtDepartment.Locked = True
    txtApprovedBy.Locked = True
    txtEmployeeName.Locked = True
    dtpBillDate.Value = Date
    dtpCoupenDate.Value = Date
    
    txtNo.Locked = False
    txtNo.Text = ""
    txtNo.Locked = True
    cmbItemList.Text = ""
    txtQty.Text = ""
    txtRate.Locked = False
    txtTotal.Locked = False
    txtRate = 0
    txtTotal = 0
    txtRate.Locked = True
    txtTotal.Locked = True
    Call mInitGrid
    Call mGetItemDetails
End Sub

Public Sub mInitGrid()
    MSHFlexGrid1.Clear
    With MSHFlexGrid1
        .Rows = 2
        .BandDisplay = flexBandDisplayHorizontal
        .ColHeaderCaption(0, 0) = "No."
        .ColWidth(0) = 500
        .ColHeaderCaption(0, 1) = "Item Description"
        .ColWidth(1) = 2400
        .ColHeaderCaption(0, 2) = "Qty"
        .ColWidth(2) = 700
        .ColHeaderCaption(0, 3) = "Rate"
        .ColWidth(3) = 800
        .ColHeaderCaption(0, 4) = "Total"
        .ColWidth(4) = 800
        .ColWidth(5) = 5

        .Row = 1
        .Col = 0
        txtNo.Left = .Left + .CellLeft
        txtNo.Width = .CellWidth - 20
        txtNo.Height = .CellHeight - 20
        txtNo.Text = 1
        .Col = 1
        cmbItemList.Left = .Left + .CellLeft
        cmbItemList.Width = .CellWidth - 20
        .Col = 2
        txtQty.Left = .Left + .CellLeft
        txtQty.Width = .CellWidth - 20
        txtQty.Height = .CellHeight - 20
        .Col = 3
        txtRate.Left = .Left + .CellLeft
        txtRate.Width = .CellWidth - 20
        txtRate.Height = .CellHeight - 20
        .Col = 4
        txtTotal.Left = .Left + .CellLeft
        txtTotal.Width = .CellWidth - 20
        txtTotal.Height = .CellHeight - 20
    End With
    
End Sub

Public Sub mAddRowToGrid()
    'assume that all control in grid located at last row
    cmbItemList.Text = ""
    txtQty = ""
    txtRate.Locked = False
    txtTotal.Locked = False
    txtRate = 0
    txtTotal = 0
    txtRate.Locked = True
    txtTotal.Locked = True
    txtNo.Locked = False
    txtNo.Text = MSHFlexGrid1.Rows
    txtNo.Locked = True
    cmbItemList.SetFocus
    
    If blGridClick Then Exit Sub
    
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
End Sub

Public Sub mGetItemDetails()
On Error GoTo FetchError
    Dim rsFetch As New ADODB.Recordset, i As Integer
    
    rsFetch.Open "select * from m_canteen_item where status=1", gStrConnectionString, adOpenKeyset, adLockOptimistic
    cmbItemList.Clear
    If rsFetch.RecordCount > 0 Then
        ReDim tpItemLst(rsFetch.RecordCount - 1)
        For i = 0 To rsFetch.RecordCount - 1
            tpItemLst(i).iItemCode = rsFetch.Fields(0)
            tpItemLst(i).strDescription = rsFetch.Fields(1)
            cmbItemList.AddItem rsFetch.Fields(1)
            tpItemLst(i).cRate = rsFetch.Fields(3)
            rsFetch.MoveNext
        Next
    End If
Exit Sub
FetchError:
    MsgBox "Error: " & Err.Number & vbCrLf & "Desc: Error in Updation of record" & vbCrLf & Err.Description, vbCritical, "Update"
    Err.Clear
End Sub

Private Sub mnuAdd_Click()
    blGridClick = False
    Call mAddRowToGrid
End Sub

Private Sub mnuRemove_Click()
    Dim iDel As Integer, i As Integer
    
    iDel = Val(InputBox("Enter Sr. No. to delete Item", "Delete", 0))
    If iDel > 0 Then
        MSHFlexGrid1.Col = 0
        For i = 1 To MSHFlexGrid1.Rows - 1
            MSHFlexGrid1.Row = i
            If iDel = Val(MSHFlexGrid1.Text) Then
                MSHFlexGrid1.RemoveItem (i - 1)
                MSHFlexGrid1.Refresh
                Exit For
            End If
        Next
    End If
    
    For i = 1 To MSHFlexGrid1.Rows - 1
        MSHFlexGrid1.Col = 0
        MSHFlexGrid1.Row = i
        MSHFlexGrid1.Text = i
    Next
    
    txtId.Locked = False
    txtId.Text = MSHFlexGrid1.Rows - 1
    txtId.Locked = True
End Sub

Private Sub MSHFlexGrid1_Click()
    blGridClick = True
    Call mReadFromGrid(MSHFlexGrid1.Row)
End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuRight
    End If
End Sub

Private Sub txtQty_Change()
If Val(txtQty) > 0 Then
    txtTotal = Val(txtQty) * Val(txtRate)
Else
    txtTotal = 0
End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call mWriteToGrid(txtNo.Text)
        blGridClick = False
        Call mAddRowToGrid
        KeyAscii = 0
    End If
End Sub

Public Sub mWriteToGrid(iRowNo As Integer)
    With MSHFlexGrid1
        .TextMatrix(iRowNo, 0) = txtNo.Text
        .TextMatrix(iRowNo, 1) = cmbItemList.Text
        .TextMatrix(iRowNo, 2) = txtQty.Text
        .TextMatrix(iRowNo, 3) = txtRate.Text
        .TextMatrix(iRowNo, 4) = txtTotal.Text
    End With
End Sub

Public Sub mReadFromGrid(iRowNo As Integer)

With MSHFlexGrid1
    txtNo.Locked = False
    txtNo.Text = .TextMatrix(iRowNo, 0)
    If txtNo.Text = "" Then txtNo.Text = iRowNo - 1
    txtNo.Locked = True
    cmbItemList.Text = .TextMatrix(iRowNo, 1)
    txtRate.Locked = False
    txtTotal.Locked = False
    txtRate.Text = .TextMatrix(iRowNo, 3)
    txtTotal.Text = .TextMatrix(iRowNo, 4)
    txtRate.Locked = True
    txtTotal.Locked = True
    txtQty.Text = .TextMatrix(iRowNo, 2)
    
    .Col = 0
    .Row = .Rows - 1
    If .Text = "" Then
        .Rows = .Rows - 1
    End If
    
End With
End Sub


