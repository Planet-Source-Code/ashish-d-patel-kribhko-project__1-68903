VERSION 5.00
Begin VB.Form frmChangePassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   2625
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   150
      TabIndex        =   2
      Top             =   210
      Width           =   4425
      Begin VB.TextBox txtNewPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   1875
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   690
         Width           =   2355
      End
      Begin VB.TextBox txtOldPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   1875
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   270
         Width           =   2355
      End
      Begin VB.TextBox txtRetypePassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   1875
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1110
         Width           =   2355
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Retype Password:"
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
         Left            =   120
         TabIndex        =   5
         Top             =   1170
         Width           =   1650
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Old Password:"
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
         Left            =   465
         TabIndex        =   4
         Top             =   330
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "New Password:"
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
         Left            =   375
         TabIndex        =   3
         Top             =   735
         Width           =   1395
      End
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
      Left            =   2430
      TabIndex        =   1
      Top             =   2070
      Width           =   1005
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Change"
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
      Left            =   1350
      TabIndex        =   0
      Top             =   2070
      Width           =   1005
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCancel_Click()
    txtOldPassword.Text = ""
    txtNewPassword.Text = ""
    txtRetypePassword.Text = ""
End Sub

Private Sub cmdChange_Click()
On Error GoTo ChangeError
    Dim rsChange As New ADODB.Recordset, cnChange As New ADODB.Connection
    
    rsChange.Open "select emp_code,emp_passwd from m_employee where emp_code='" & gStrUser & "'" _
                , gStrConnectionString, adOpenKeyset, adLockOptimistic
    
    If Len(txtOldPassword) = 0 Then
        MsgBox "Password can not be left blank", vbInformation, "Change Password"
        txtOldPassword.SetFocus
        Exit Sub
    ElseIf txtOldPassword.Text <> rsChange.Fields(1) Then
        MsgBox "Incorrect Old Password", vbInformation, "Change Password"
        txtOldPassword.Text = ""
        txtOldPassword.SetFocus
        Exit Sub
    ElseIf Len(txtNewPassword) = 0 Then
        MsgBox "New Password can not be left blank", vbInformation, "Change Password"
        txtNewPassword.SetFocus
        Exit Sub
    ElseIf txtNewPassword.Text <> txtRetypePassword.Text Then
        MsgBox "New Password and Retype Password doesn't match", vbInformation, "Change Password"
        Exit Sub
    End If
    
    rsChange.Close
    cnChange.ConnectionString = gStrConnectionString
    cnChange.Open
    cnChange.Execute "update m_employee set emp_passwd='" & txtNewPassword.Text & "' where " & _
                    " emp_code='" & gStrUser & "'"
    MsgBox "Password Change Successfully", vbInformation, "Change Password"
    Unload Me
Exit Sub
ChangeError:
    MsgBox "Error :" & Err.Number & vbCrLf & "Desc: " & Err.Description, vbCritical, "Password Change"
    Err.Clear
End Sub

Private Sub Form_Load()
    Call gFormCenter(Me)
    Call cmdCancel_Click
End Sub
