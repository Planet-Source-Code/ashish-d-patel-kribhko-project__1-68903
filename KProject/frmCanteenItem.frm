VERSION 5.00
Begin VB.Form frmCanteenItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Canteen Item"
   ClientHeight    =   3435
   ClientLeft      =   2445
   ClientTop       =   3435
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Canteen Item Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2385
      Left            =   210
      TabIndex        =   0
      Top             =   180
      Width           =   5415
      Begin VB.CommandButton cmdItemTypeList 
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
         TabIndex        =   10
         Top             =   1350
         Width           =   735
      End
      Begin VB.TextBox txtName 
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
         Left            =   1170
         TabIndex        =   4
         Top             =   915
         Width           =   4035
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
         Left            =   1170
         TabIndex        =   3
         Top             =   480
         Width           =   1245
      End
      Begin VB.TextBox txtItemType 
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
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1350
         Width           =   2025
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1170
         TabIndex        =   1
         Top             =   1800
         Width           =   1275
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Rate:"
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
         Left            =   570
         TabIndex        =   8
         Top             =   1860
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Type:"
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
         Left            =   555
         TabIndex        =   7
         Top             =   1380
         Width           =   525
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Name:"
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
         Left            =   480
         TabIndex        =   6
         Top             =   960
         Width           =   600
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
         Left            =   855
         TabIndex        =   5
         Top             =   540
         Width           =   240
      End
   End
   Begin KProject.ActionButton ActionButton1 
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   2760
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
Attribute VB_Name = "frmCanteenItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim iItemTypeId As Integer


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
    frmSelect.strSql = "select i_id as ID ,i_name as Item, it_name as [Item Type] from m_canteen_item mi, m_canteen_itemtype mt" & _
                        " where mi.status=1 and mi.it_id=mt.it_id"
    gBlCanteenItem = True
    frmSelect.Show vbModal
    
    If gIntCanteenItem > 0 Then
        Dim rsDisplay As New ADODB.Recordset
        
        rsDisplay.Open "select i_id ,i_name, it_name,i_rate,mi.it_id from m_canteen_item mi, m_canteen_itemtype mt" & _
                        " where mi.status=1 and mi.it_id=mt.it_id and i_id=" & gIntCanteenItem, gStrConnectionString, adOpenKeyset, adLockOptimistic
        txtId = rsDisplay.Fields(0)
        txtName = rsDisplay.Fields(1)
        txtItemType = rsDisplay.Fields(2)
        txtRate = rsDisplay.Fields(3)
        iItemTypeId = gIntCanteenItem
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
        MsgBox "Item Name can not be left blank" & vbCrLf & "Sub department can not exist without Department", vbInformation, "Update"
        txtName.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtItemType)) = 0 Then
        MsgBox "Please select Item Type", vbInformation, "Update"
        cmdItemTypeList.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtRate)) = 0 Then
        MsgBox "Rate can not be left blank", vbInformation, "Update"
        txtRate.SetFocus
        Exit Sub
    End If

    ActionButton1.blSave = True
    ActionButton1.SaveSql = "Insert into m_canteen_item values (" & txtId & ",'" & txtName & "'," & iItemTypeId & "," & Val(txtRate) & ",1)"
    Call mResetControl(False)
End Sub

Private Sub cmdItemTypeList_Click()
    gBlListCanteenItemType = True
    frmList.strSql = "select it_id as ID,it_name as [Item Type] from m_canteen_itemtype where status=1"
    frmList.Show vbModal
    txtItemType.Locked = False
    txtItemType = gstrListCanteenItemType
    txtItemType.Locked = True
    iItemTypeId = gintListCanteenItemType
End Sub

Private Sub Form_Load()

    Call gFormCenter(Me)
    With ActionButton1
        .ConnectionString = gStrConnectionString
        .PrimaryKeyField = "i_id"
        .Table = "m_canteen_item"
    End With
    Call mResetControl(False)
End Sub

Public Sub mResetControl(ByVal blEnable As Boolean)
    txtName.Enabled = blEnable
    txtItemType.Enabled = blEnable
    txtRate.Enabled = blEnable
End Sub

Public Sub mClearControl()
    txtId.Text = ""
    txtName.Text = ""
    txtItemType = ""
    txtRate = ""
End Sub
