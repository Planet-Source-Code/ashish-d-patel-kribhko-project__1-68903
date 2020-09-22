VERSION 5.00
Begin VB.Form frmDesignationMaster 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Designation Master"
   ClientHeight    =   3195
   ClientLeft      =   2025
   ClientTop       =   2520
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Designation Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   450
      TabIndex        =   0
      Top             =   330
      Width           =   5145
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
         Left            =   1515
         TabIndex        =   2
         Top             =   900
         Width           =   3045
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
         Left            =   1515
         TabIndex        =   1
         Top             =   480
         Width           =   1245
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
         Left            =   810
         TabIndex        =   4
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
         Left            =   1185
         TabIndex        =   3
         Top             =   540
         Width           =   240
      End
   End
   Begin KProject.ActionButton ActionButton1 
      Height          =   495
      Left            =   420
      TabIndex        =   5
      Top             =   2100
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
Attribute VB_Name = "frmDesignationMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

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
    frmSelect.strSql = "select desg_id as ID ,desg_name as Designation from m_Designation" & _
                        " where status=1"
    gBlDesignation = True
    frmSelect.Show vbModal
    
    If gIntDesignation > 0 Then
        Dim rsDisplay As New ADODB.Recordset
        
        rsDisplay.Open "select desg_id,desg_name from m_Designation where desg_id=" & gIntDesignation, _
                        gStrConnectionString, adOpenKeyset, adLockOptimistic
        txtId = rsDisplay.Fields(0)
        txtName = rsDisplay.Fields(1)
        ActionButton1.blModify = True
        ActionButton1.iModifyRecord = txtId
        Call mResetControl(True)
    Else
        ActionButton1.blModify = False
        ActionButton1.blSave = False
    End If
End Sub

Private Sub ActionButton1_UpdateClick()
    
    If Len(Trim(txtName)) = 0 Then
        MsgBox "Designation Name can not be left blank", vbInformation, "Update"
        Exit Sub
    End If

    ActionButton1.blSave = True
    ActionButton1.SaveSql = "Insert into m_Designation values (" & txtId & ",'" & txtName & "',1)"
    Call mResetControl(False)
End Sub

Private Sub Form_Load()

    Call gFormCenter(Me)
    With ActionButton1
        .ConnectionString = gStrConnectionString
        .PrimaryKeyField = "desg_id"
        .Table = "m_Designation"
    End With
    Call mResetControl(False)
End Sub

Public Sub mResetControl(ByVal blEnable As Boolean)
    txtName.Enabled = blEnable
End Sub

Public Sub mClearControl()
    txtId.Text = ""
    txtName.Text = ""
End Sub


