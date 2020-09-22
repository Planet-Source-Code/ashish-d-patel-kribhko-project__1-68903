VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTimeCard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Time Card"
   ClientHeight    =   6495
   ClientLeft      =   2700
   ClientTop       =   1290
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
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
      Left            =   3690
      TabIndex        =   2
      Top             =   240
      Width           =   1005
   End
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
      Height          =   5445
      Left            =   180
      TabIndex        =   0
      Top             =   900
      Width           =   5085
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   4995
         Left            =   120
         TabIndex        =   1
         Top             =   330
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   8811
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   360
      Left            =   2010
      TabIndex        =   3
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
      Format          =   19791875
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
      Left            =   720
      TabIndex        =   4
      Top             =   300
      Width           =   1200
   End
End
Attribute VB_Name = "frmTimeCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdView_Click()
On Error GoTo ViewError
    Dim rsSchedule As New ADODB.Recordset
    Dim strSql As String

    strSql = "select ea_date,ea_in_time,ea_out_time from t_employee_attendence where emp_code='" & gStrUser & _
        "' and month(ea_date)=" & Month(DTPicker1) & " and year(ea_date)= " & Year(DTPicker1)
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
    .ColWidth(0) = 1500
    .ColWidth(1) = 1500
    .ColWidth(2) = 1500
    .Row = 0
    .Col = 0
    .Text = "Date"
    .Col = 1
    .Text = "In Time"
    .Col = 2
    .Text = "Out Time"
End With
End Sub

Private Sub Form_Load()
    Call gFormCenter(Me)
End Sub


