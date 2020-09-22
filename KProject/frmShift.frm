VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmShift 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Shift Master"
   ClientHeight    =   4260
   ClientLeft      =   2385
   ClientTop       =   2550
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Shift Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2745
      Left            =   210
      TabIndex        =   7
      Top             =   330
      Width           =   5595
      Begin VB.TextBox txtEarlyGo 
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
         Left            =   3780
         MaxLength       =   2
         TabIndex        =   5
         Top             =   1740
         Width           =   555
      End
      Begin VB.TextBox txtLateAllow 
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
         Left            =   1140
         MaxLength       =   2
         TabIndex        =   4
         Top             =   1740
         Width           =   555
      End
      Begin MSComCtl2.DTPicker dtpEndTime 
         Height          =   360
         Left            =   3750
         TabIndex        =   3
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   19726338
         UpDown          =   -1  'True
         CurrentDate     =   38818
      End
      Begin VB.TextBox txtRemark 
         Height          =   360
         Left            =   1140
         TabIndex        =   6
         Top             =   2160
         Width           =   4035
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
         Left            =   1155
         TabIndex        =   1
         Top             =   900
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
         Left            =   1155
         TabIndex        =   8
         Top             =   480
         Width           =   1245
      End
      Begin MSComCtl2.DTPicker dtpStartTime 
         Height          =   360
         Left            =   1140
         TabIndex        =   2
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   19726338
         UpDown          =   -1  'True
         CurrentDate     =   38818
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Min."
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
         Left            =   4410
         TabIndex        =   17
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Min."
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
         Left            =   1800
         TabIndex        =   16
         Top             =   1800
         Width           =   360
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
         Index           =   5
         Left            =   285
         TabIndex        =   15
         Top             =   2190
         Width           =   765
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Early go:"
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
         Index           =   4
         Left            =   2835
         TabIndex        =   14
         Top             =   1770
         Width           =   795
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Late Allow:"
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
         Left            =   105
         TabIndex        =   13
         Top             =   1770
         Width           =   960
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "End Time:"
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
         Left            =   2730
         TabIndex        =   12
         Top             =   1380
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Start Time:"
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
         Left            =   120
         TabIndex        =   11
         Top             =   1350
         Width           =   960
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   540
         Width           =   240
      End
   End
   Begin KProject.ActionButton ActionButton1 
      Height          =   495
      Left            =   420
      TabIndex        =   0
      Top             =   3300
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
Attribute VB_Name = "frmShift"
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
    frmSelect.strSql = "select shf_id as ID ,shf_name as Shiftname,shf_start_time as StartTime from m_Shift" & _
                        " where status=1"
    gBlShift = True
    frmSelect.Show vbModal
    
    If gIntShift > 0 Then
        Dim rsDisplay As New ADODB.Recordset
        
        rsDisplay.Open "select * from m_shift where shf_id=" & gIntShift, _
                        gStrConnectionString, adOpenKeyset, adLockOptimistic
        txtId = rsDisplay.Fields(0)
        txtName = rsDisplay.Fields(1)
        dtpStartTime.Value = rsDisplay.Fields(2)
        dtpEndTime.Value = rsDisplay.Fields(3)
        txtEarlyGo.Text = rsDisplay.Fields(5)
        txtLateAllow.Text = rsDisplay.Fields(4)
        txtRemark.Text = rsDisplay(6)
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
        MsgBox "Shift Name can not be left blank", vbInformation, "Update"
        Exit Sub
    ElseIf Val(Trim(txtEarlyGo)) > 60 Then
        MsgBox "Early go minutes must be less than 60", vbInformation, "Update"
        Exit Sub
    ElseIf Val(Trim(txtLateAllow)) > 60 Then
        MsgBox "Late Allowed minutes must be less than 60", vbInformation, "Update"
        Exit Sub
    End If

    ActionButton1.blSave = True
    ActionButton1.SaveSql = "Insert into m_Shift values (" & txtId.Text & ",'" & txtName.Text & "',#" & _
                             dtpStartTime.Value & "#,#" & dtpEndTime.Value & "#," & Val(txtLateAllow) & _
                             "," & Val(txtEarlyGo) & ",'" & txtRemark & "',1)"
    Call mResetControl(False)
End Sub

Private Sub Form_Load()

    Call gFormCenter(Me)
    With ActionButton1
        .ConnectionString = gStrConnectionString
        .PrimaryKeyField = "Shf_id"
        .Table = "m_Shift"
    End With
    Call mResetControl(False)
End Sub

Public Sub mResetControl(ByVal blEnable As Boolean)
    txtName.Enabled = blEnable
    dtpStartTime.Enabled = blEnable
    dtpEndTime.Enabled = blEnable
    txtLateAllow.Enabled = blEnable
    txtEarlyGo.Enabled = blEnable
    txtRemark.Enabled = blEnable
End Sub

Public Sub mClearControl()
    txtId.Text = ""
    txtName.Text = ""
    txtRemark = ""
    dtpStartTime.Value = "12:00:00"
    dtpEndTime.Value = "12:00:00"
    txtLateAllow = ""
    txtEarlyGo = ""
End Sub


