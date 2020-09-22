VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   3240
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   150
      TabIndex        =   6
      Top             =   660
      Width           =   4275
      Begin VB.TextBox txtPassword 
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
         Left            =   1740
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   690
         Width           =   2325
      End
      Begin VB.TextBox txtUserName 
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
         Left            =   1740
         TabIndex        =   0
         Top             =   270
         Width           =   2325
      End
      Begin VB.Frame Frame2 
         Caption         =   "Purpose of login"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   0
         TabIndex        =   9
         Top             =   1230
         Width           =   4275
         Begin VB.OptionButton optAttendence 
            Caption         =   "&Attendence (In and Out )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Value           =   -1  'True
            Width           =   2445
         End
         Begin VB.OptionButton optKsystem 
            Caption         =   "K &System"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2730
            TabIndex        =   3
            Top             =   240
            Width           =   1305
         End
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Password:"
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
         Left            =   705
         TabIndex        =   8
         Top             =   705
         Width           =   945
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Employee Code:"
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
         Left            =   135
         TabIndex        =   7
         Top             =   345
         Width           =   1515
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   330
      Top             =   2760
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "O&K"
      Default         =   -1  'True
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
      Left            =   1140
      TabIndex        =   4
      Top             =   2700
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
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
      Left            =   2370
      TabIndex        =   5
      Top             =   2700
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   180
      TabIndex        =   10
      Top             =   60
      Width           =   4245
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
    txtUserName.Text = ""
    txtPassword = ""
    txtUserName.SetFocus
End Sub

Private Sub cmdOK_Click()
On Error GoTo LoginError

    Dim rsCheck As New ADODB.Recordset, strSql As String, strName As String
    Dim cnCheck As New ADODB.Connection
    
    strSql = "select * from m_employee where emp_code='" & txtUserName & "' and emp_passwd='" & _
                txtPassword & "'"
    rsCheck.Open strSql, gStrConnectionString, adOpenKeyset, adLockOptimistic
    
    If rsCheck.RecordCount > 0 And optAttendence Then
        strName = rsCheck.Fields("emp_salutation") & " " & rsCheck.Fields("emp_fname")
        strSql = "select * from t_employee_attendence where emp_code='" & txtUserName & "' and " & _
            " ea_date=#" & Date & "#"
        
        rsCheck.Close
        rsCheck.Open strSql, gStrConnectionString, adOpenKeyset, adLockOptimistic
        cnCheck.ConnectionString = gStrConnectionString
        cnCheck.Open
        If rsCheck.RecordCount > 0 Then
            rsCheck.Close
            strSql = "select * from t_employee_attendence where emp_code='" & txtUserName & "' and " & _
            " ea_date=#" & Date & "# and ea_out_time <> null"
            rsCheck.Open strSql, gStrConnectionString, adOpenKeyset, adLockOptimistic
                If rsCheck.RecordCount > 0 Then
                    MsgBox "You have already entered OUT TIME", vbInformation, "Attendence"
                    cmdCancel_Click
                    Exit Sub
                End If
            strSql = "Update t_employee_attendence set ea_out_time=#" & Format(Now, "HH:MM:SS") & _
                    "# where emp_code='" & txtUserName & "' and ea_date=#" & Date & "#"
            cnCheck.Execute strSql
            MsgBox "Goodbye " & strName & ", your OUT TIME is updated" & vbCrLf & vbCrLf & _
                "NOTE: In case of wrong message Contact your HOD", vbInformation, "Attendence"
        Else
            strSql = "select max(ea_id) from t_employee_attendence"
            rsCheck.Close
            rsCheck.Open strSql, gStrConnectionString, adOpenKeyset, adLockOptimistic
            strSql = "Insert into t_employee_attendence(ea_id,ea_date,emp_code,ea_in_time) values (" & _
                rsCheck.Fields(0) + 1 & ",#" & Date & "#,'" & txtUserName & "',#" & Format(Now, "HH:MM:SS") & "#)"
            Debug.Print strSql
            cnCheck.Execute strSql
            MsgBox "Welcome " & strName & ", your IN TIME is updated " & vbCrLf & vbCrLf & _
                "NOTE: In case of wrong message Contact your HOD", vbInformation, "Attendence"
        End If
        cmdCancel_Click
    ElseIf rsCheck.RecordCount > 0 And optKsystem Then
        gStrUser = txtUserName
        gIntUserId = rsCheck.Fields("emp_id")
        Unload Me
        frmMain.Show
    Else
        MsgBox "Wrong Employee code or password", vbExclamation, "Login Check"
        txtUserName.SetFocus
    End If
    
Exit Sub
LoginError:
    MsgBox "Error: " & Err.Number & vbCrLf & "Desc :" & Err.Description, vbCritical, "Login Error"
    Err.Clear
End Sub

Private Sub Form_Load()
    Me.Show
    Call gFormCenter(Me)
    If Strings.StrComp(CStr(Date), Format(Date, "dd/MMM/yyyy"), vbTextCompare) <> 0 Then
        MsgBox ("Please set the REGIONAL OPTIONS's date format to 'dd/MMM/yyyy'")
        End
    End If
    gStrConnectionString = "Provider=Microsoft.Jet.Oledb.4.0;Data Source=" & App.Path & "\KSystemdb.mdb"
End Sub

Private Sub Timer1_Timer()
    Label1.Caption = "Time : " & Format(Now, "HH:MM:SS")
End Sub
