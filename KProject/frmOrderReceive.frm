VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmOrderReceive 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Order Receive"
   ClientHeight    =   6615
   ClientLeft      =   2130
   ClientTop       =   1410
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
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
      Left            =   3000
      TabIndex        =   15
      Top             =   5970
      Width           =   1005
   End
   Begin VB.CommandButton cmdReceived 
      Caption         =   "&Receive"
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
      Left            =   1950
      TabIndex        =   14
      Top             =   5970
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Caption         =   "Order"
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
      TabIndex        =   0
      Top             =   150
      Width           =   5835
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
         TabIndex        =   7
         Top             =   1560
         Width           =   3495
      End
      Begin VB.CommandButton cmdOrderCodeList 
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
         Left            =   3090
         TabIndex        =   6
         Top             =   720
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
         TabIndex        =   5
         Top             =   1140
         Width           =   3495
      End
      Begin VB.TextBox txtOrderCode 
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
         TabIndex        =   4
         Top             =   720
         Width           =   1635
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
         Left            =   4440
         TabIndex        =   3
         Top             =   300
         Visible         =   0   'False
         Width           =   1245
      End
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
         TabIndex        =   1
         Top             =   2040
         Width           =   5655
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   3045
            Left            =   120
            TabIndex        =   2
            Top             =   330
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   5371
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
      Begin MSComCtl2.DTPicker dtpOrderDate 
         Height          =   360
         Left            =   1380
         TabIndex        =   8
         Top             =   300
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
         TabIndex        =   13
         Top             =   1590
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
         TabIndex        =   12
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Receive Date:"
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
         Left            =   30
         TabIndex        =   11
         Top             =   390
         Width           =   1290
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
         TabIndex        =   10
         Top             =   780
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
         Left            =   4140
         TabIndex        =   9
         Top             =   390
         Visible         =   0   'False
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmOrderReceive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCancel_Click()
    txtOrderCode.Locked = False
    txtDepartment.Locked = False
    txtApprovedBy.Locked = False
    txtOrderCode.Text = ""
    txtDepartment.Text = ""
    txtApprovedBy.Text = ""
    txtOrderCode.Locked = True
    txtDepartment.Locked = True
    txtApprovedBy.Locked = True
    MSHFlexGrid1.Clear
    Call mInitGrid
End Sub

Private Sub cmdOrderCodeList_Click()
On Error GoTo OrderListError
    
    Dim rsOrder As New ADODB.Recordset
    
    MSHFlexGrid1.Clear
    MSHFlexGrid1.Rows = 2

    gBlListOrder = True
    frmList.strSql = "select re_date as [Order Date], re_order_id as [Order Code] " & _
                        " , dept_name as Department,re_id from view_orderlist where re_order_status='No'"
    frmList.Show vbModal
    txtOrderCode.Locked = False
    txtOrderCode = gstrListOrder
    txtOrderCode.Locked = True
    rsOrder.Open "Select dept_name,emp_fname from view_orderlist where re_order_id='" & txtOrderCode.Text & "'" _
                    , gStrConnectionString, adOpenKeyset, adLockOptimistic
    txtDepartment.Locked = False
    txtApprovedBy.Locked = False
    txtDepartment.Text = rsOrder.Fields(0)
    txtApprovedBy.Text = rsOrder.Fields(1)
    txtDepartment.Locked = True
    txtApprovedBy.Locked = True
    rsOrder.Close
    rsOrder.Open "Select red_id ,prt_name,red_qty,prt_price,red_qty*prt_price as Total from " & _
                "view_Orderdetail where re_id=" & gintListOrder, gStrConnectionString, adOpenKeyset, adLockOptimistic
    Set MSHFlexGrid1.Recordset = rsOrder
    
    Call mInitGrid
    Exit Sub
OrderListError:
    MsgBox "Error: " & Err.Number & vbCrLf & "Desc :" & Err.Description, vbCritical, "OrderList Error"
    Err.Clear
End Sub

Public Sub mInitGrid()
    
    With MSHFlexGrid1
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
    End With
End Sub

Private Sub cmdReceived_Click()
On Error GoTo ReceiveError
    Dim cnReceive As New ADODB.Connection
    Dim strQ As String
    
    cnReceive.ConnectionString = gStrConnectionString
    cnReceive.Open
    strQ = "Update t_requirement_entry set re_order_status='Yes',re_order_recv_date=#" & dtpOrderDate.Value & _
            "# where re_id=" & gintListOrder
    cnReceive.Execute strQ
    MsgBox "Order details updated ", vbInformation, "Update"
    cmdCancel_Click
Exit Sub
ReceiveError:
    MsgBox "Error :" & Err.Number & vbCrLf & "Desc :" & Err.Description, vbCritical, "Receive Error"
    Err.Clear
End Sub

Private Sub Form_Load()
    Call gFormCenter(Me)
End Sub
