VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRequirement 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Requirement Entry"
   ClientHeight    =   6630
   ClientLeft      =   2280
   ClientTop       =   1485
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   30
      Top             =   5370
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Caption         =   "Requirement"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   150
      TabIndex        =   10
      Top             =   180
      Width           =   5835
      Begin VB.Frame Frame2 
         Caption         =   "Order Details "
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
         TabIndex        =   19
         Top             =   2040
         Width           =   5655
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
            TabIndex        =   8
            Top             =   450
            Width           =   675
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
            TabIndex        =   7
            Top             =   450
            Width           =   825
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
            TabIndex        =   6
            Top             =   450
            Width           =   735
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
            TabIndex        =   9
            Top             =   450
            Width           =   615
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   2595
            Left            =   120
            TabIndex        =   20
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
         TabIndex        =   13
         Top             =   300
         Width           =   1245
      End
      Begin VB.TextBox txtOrderCode 
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
         TabIndex        =   1
         Top             =   720
         Width           =   1245
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
         TabIndex        =   12
         Top             =   1140
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
         TabIndex        =   3
         Top             =   1140
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
         TabIndex        =   11
         Top             =   1560
         Width           =   3495
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
         TabIndex        =   4
         Top             =   1560
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtpOrderDate 
         Height          =   360
         Left            =   3990
         TabIndex        =   2
         Top             =   720
         Width           =   1665
         _ExtentX        =   2937
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
         CurrentDate     =   38823
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
         Left            =   1080
         TabIndex        =   18
         Top             =   390
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Order Code:"
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
         TabIndex        =   17
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Order Date:"
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
         Left            =   2880
         TabIndex        =   16
         Top             =   780
         Width           =   1035
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
         TabIndex        =   15
         Top             =   1200
         Width           =   1095
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
         TabIndex        =   14
         Top             =   1590
         Width           =   1230
      End
   End
   Begin KProject.ActionButton ActionButton1 
      Height          =   495
      Left            =   420
      TabIndex        =   0
      Top             =   6000
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
   Begin VB.Menu mnuRight 
      Caption         =   "Right"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "&Remove"
      End
   End
End
Attribute VB_Name = "frmRequirement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim iRow As Integer, blGridClick As Boolean, blMenuAdd As Boolean
Dim iStartTop As Integer
Private Type ItemList
    iItemCode As Integer
    strDescription As String
    cRate As Currency
End Type
Dim tpItemLst() As ItemList
Private Sub cmbItemList_Click()
'Dim tempObject As New clsItemList
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

Private Sub cmdApprovedBy_Click()
    gBlListEmployee = True
    frmList.strSql = "select emp_id as ID,emp_fname as [First Name],emp_lname as [Last Name],dept_name as Department from view_employee where status=1"
    frmList.Show vbModal
    txtApprovedBy.Locked = False
    txtApprovedBy = gstrListEmployee
    txtApprovedBy.Locked = True
    gBlListEmployee = False
End Sub

Private Sub cmdDepartmentList_Click()
    gBlListDept = True
    frmList.strSql = "select dept_id as ID,dept_name as Department from m_department where status=1"
    frmList.Show vbModal
    txtDepartment.Locked = False
    txtDepartment = gstrListDeptName
    txtDepartment.Locked = True
End Sub

Private Sub Form_Load()
    
    Call gFormCenter(Me)
    
    With ActionButton1
        .ConnectionString = gStrConnectionString
        .PrimaryKeyField = "re_id"
        .Table = "t_requirement_entry"
    End With
    
    Call mGetItemDetails
    Call mResetControl(False)
    Call mInitGrid
    iStartTop = txtNo.Top
End Sub

Private Sub ActionButton1_AddClick(iNextNo As Integer)

    Call mClearControl
    txtId = iNextNo + 1
    Call mResetControl(True)
    txtOrderCode.SetFocus
End Sub

Private Sub ActionButton1_CancelClick()
    Call mClearControl
    Call mResetControl(False)
End Sub

Private Sub ActionButton1_DeleteClick()
    With ActionButton1
        .iModifyRecord = txtId
        .strDelete = txtOrderCode
    End With
    Call mClearControl
End Sub

Private Sub ActionButton1_ModifyClick()
    frmSelect.strSql = "select re_id as ID ,re_order_id as [Order Code], re_date  as [Order Date] from" & _
                        " t_requirement_entry where status=1"
    gBlRequirement = True
    frmSelect.Show vbModal
    
    Frame2.Caption = "Order Details ( To Add/Remove Row right click on grid )"
    
    If gIntRequirement > 0 Then
        Dim rsDisplay As New ADODB.Recordset
        
        rsDisplay.Open "select * from view_requirement where re_id=" & gIntRequirement, _
                        gStrConnectionString, adOpenKeyset, adLockOptimistic
        txtId = rsDisplay.Fields(0)
        txtOrderCode.Text = rsDisplay.Fields(1)
        txtDepartment.Text = rsDisplay.Fields(2)
        dtpOrderDate.Value = rsDisplay.Fields(3)
        txtApprovedBy.Text = rsDisplay.Fields(4)
        gintListDeptId = rsDisplay.Fields("dept_id")
        gintListEmployee = rsDisplay.Fields("appv_emp_id")
        
        Dim rsDetail As New ADODB.Recordset, strSQ As String
        
        strSQ = " SELECT t_Requirement_Entry_Detail.Red_Id, m_Product.Prt_Name, " & _
        "t_Requirement_Entry_Detail.Red_Qty, m_Product.Prt_Price, [red_qty]*[prt_price] " & _
        " AS Total, t_Requirement_Entry_Detail.Prt_Id " & _
        " FROM m_Product INNER JOIN t_Requirement_Entry_Detail ON m_Product.Prt_id " & _
        " = t_Requirement_Entry_Detail.Prt_Id where t_requirement_entry_detail.re_id= " & txtId

        
        rsDetail.Open strSQ, gStrConnectionString, adOpenKeyset, adLockOptimistic
        
        Call mInitGrid
        With MSHFlexGrid1
            Set .Recordset = rsDetail
            Call mReadFromGrid(1)
        End With
        
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
On Error GoTo UpdateError
    Dim cnDetail As New ADODB.Connection, strQ As String, i As Integer
        
    Frame2.Caption = "Order Details"
        
    If Len(Trim(txtOrderCode)) = 0 Then
        MsgBox "Order code can not be left blank", vbInformation, "Update"
        txtOrderCode.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtDepartment)) = 0 Then
        MsgBox "Please select department", vbInformation, "Update"
        cmdDepartmentList.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtApprovedBy)) = 0 Then
        MsgBox "Please select who approve this ", vbInformation, "Update"
        cmdApprovedBy.SetFocus
        Exit Sub
    End If
    
    cnDetail.ConnectionString = gStrConnectionString
    cnDetail.Open
    cnDetail.BeginTrans

    If ActionButton1.blModify Then
        cnDetail.Execute "Delete from t_requirement_entry_detail where re_id=" & txtId
    End If
    
    With MSHFlexGrid1
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 2)) <> 0 Then
                strQ = "Insert into t_requirement_entry_detail values (" & .TextMatrix(i, 0) & _
                "," & txtId.Text & "," & .TextMatrix(i, 5) & "," & .TextMatrix(i, 2) & ",1)"
                cnDetail.Execute strQ
            End If
        Next
    End With
    
    ActionButton1.blSave = True
    ActionButton1.SaveSql = "Insert into t_requirement_entry values (" & txtId & _
                    ",'" & txtOrderCode & "'," & gintListDeptId & ",#" & dtpOrderDate.Value & _
                    "#," & gintListEmployee & ",'No',#" & Date & "#,1)"
    cnDetail.CommitTrans
    
    CrystalReport1.DataFiles(0) = App.Path & "\ksystemdb.mdb"
    CrystalReport1.ReportFileName = App.Path & "\report\rptOrderdetails.rpt"
    CrystalReport1.ReportTitle = "Order Details"
    CrystalReport1.PrintReport
    CrystalReport1.WindowState = crptMaximized
    Call mResetControl(False)
    Exit Sub
UpdateError:
    MsgBox "Error: " & Err.Number & vbCrLf & "Desc: Error in Updation of record" & vbCrLf & Err.Description, vbCritical, "Update"
    Err.Clear
End Sub


Public Sub mResetControl(ByVal blEnable As Boolean)
    txtId.Enabled = blEnable
    txtOrderCode.Enabled = blEnable
    dtpOrderDate.Enabled = blEnable
    txtDepartment.Enabled = blEnable
    txtApprovedBy.Enabled = blEnable
    cmdDepartmentList.Enabled = blEnable
    cmdApprovedBy.Enabled = blEnable
    MSHFlexGrid1.Enabled = blEnable
    txtNo.Enabled = blEnable
    cmbItemList.Enabled = blEnable
    txtQty.Enabled = blEnable
    txtRate.Enabled = blEnable
    txtTotal.Enabled = blEnable
End Sub

Public Sub mClearControl()
    txtId.Locked = False
    txtId.Text = ""
    txtId.Locked = True
    txtOrderCode.Text = ""
    txtDepartment.Locked = False
    txtDepartment.Text = ""
    txtDepartment.Locked = True
    txtApprovedBy.Locked = False
    txtApprovedBy.Text = ""
    txtApprovedBy.Locked = True
    dtpOrderDate.Value = Date
    
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
    
    rsFetch.Open "select * from m_Product where status=1", gStrConnectionString, adOpenKeyset, adLockOptimistic
    cmbItemList.Clear
    If rsFetch.RecordCount > 0 Then
        ReDim tpItemLst(rsFetch.RecordCount - 1)
        For i = 0 To rsFetch.RecordCount - 1
            tpItemLst(i).iItemCode = rsFetch.Fields(0)
            tpItemLst(i).strDescription = rsFetch(2) & " : " & rsFetch.Fields(1)
            cmbItemList.AddItem rsFetch(2) & " : " & rsFetch.Fields(1)
            tpItemLst(i).cRate = rsFetch.Fields(4)
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
End Sub

Private Sub MSHFlexGrid1_Click()
    blGridClick = True
    Call mReadFromGrid(MSHFlexGrid1.Row)
End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And ActionButton1.blModify Then
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
