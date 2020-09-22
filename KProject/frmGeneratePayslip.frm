VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmGeneratePayslip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Payslip Generate"
   ClientHeight    =   1935
   ClientLeft      =   2835
   ClientTop       =   3750
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Left            =   5460
      Top             =   1260
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   1335
      Left            =   450
      TabIndex        =   0
      Top             =   150
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   2355
      _Version        =   393216
      FullWidth       =   85
      FullHeight      =   89
   End
   Begin VB.Label Label1 
      Caption         =   "Processing Salary, it will take few minutes..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   660
      TabIndex        =   1
      Top             =   1500
      Width           =   4785
   End
End
Attribute VB_Name = "frmGeneratePayslip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public strMonth As String

Private Sub Form_Activate()
    Animation1.Play
    
    If MsgBox("Would you like to generate Payslip for " & strMonth, vbYesNo + vbQuestion, "Generate Payslip") = vbYes Then
        Timer1.Interval = 1000
        Call gCalculateSalary
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    If gCheckSalaryCalculation Then Unload Me

    Animation1.Open App.Path & "\working.avi"
End Sub

Private Sub Timer1_Timer()
    If gBlPayslipGenComplete And gCheckSalaryCalculation Then
        Animation1.Stop
        Timer1.Interval = 0
        Unload Me
    End If
End Sub
