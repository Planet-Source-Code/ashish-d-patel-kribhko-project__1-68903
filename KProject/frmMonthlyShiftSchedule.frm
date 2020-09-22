VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMonthlyShiftSchedule 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Monthly Shift Schedule"
   ClientHeight    =   4365
   ClientLeft      =   2265
   ClientTop       =   1935
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3225
      Left            =   150
      TabIndex        =   3
      Top             =   990
      Width           =   6915
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2745
         Left            =   210
         TabIndex        =   4
         Top             =   330
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   4842
         _Version        =   393216
         FixedCols       =   0
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
         _Band(0).Cols   =   2
      End
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
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
      Left            =   4320
      TabIndex        =   2
      Top             =   240
      Width           =   1005
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   360
      Left            =   2640
      TabIndex        =   1
      Top             =   240
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
      CustomFormat    =   "MMM-yyyy"
      Format          =   53608451
      CurrentDate     =   38833
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Select Month:"
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
      Left            =   1350
      TabIndex        =   0
      Top             =   300
      Width           =   1200
   End
End
Attribute VB_Name = "frmMonthlyShiftSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdView_Click()
On Error GoTo ViewError
    Dim rsSchedule As New ADODB.Recordset
    Dim strSql As String

    strSql = "select * from view_employeemonthlyshiftschedule where emp_id=" & gIntUserId & _
        " and month([Start date])=" & Month(DTPicker1) & " and year([Start date])=" & Year(DTPicker1)
    rsSchedule.Open strSql, gStrConnectionString, adOpenKeyset, adLockOptimistic
    
    If rsSchedule.RecordCount > 0 Then
        Set MSHFlexGrid1.Recordset = rsSchedule
    Else
        MsgBox "Record Not found ", vbInformation, "View Schedule"
        MSHFlexGrid1.Clear
    End If
    Call mFormatGrid
Exit Sub
ViewError:
    MsgBox "Error :" & Err.Number & vbCrLf & "Desc :" & Err.Description, vbCritical, "View Error"
    Err.Clear
End Sub

Public Sub mFormatGrid()
With MSHFlexGrid1
    .ColWidth(0) = 1200
    .ColWidth(1) = 1260
    .ColWidth(2) = 1260
    .ColWidth(3) = 1260
    .ColWidth(4) = 1260
    .ColWidth(5) = 5
End With
End Sub

Private Sub Form_Load()
    Call gFormCenter(Me)
End Sub
