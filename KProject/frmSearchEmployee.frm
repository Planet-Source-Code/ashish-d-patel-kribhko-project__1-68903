VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSearchEmployee 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Employee"
   ClientHeight    =   7200
   ClientLeft      =   1335
   ClientTop       =   555
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Specify Criteria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2625
      Left            =   240
      TabIndex        =   4
      Top             =   210
      Width           =   7905
      Begin VB.TextBox txtDesignationCr 
         Alignment       =   2  'Center
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
         Left            =   2820
         TabIndex        =   19
         Text            =   "="
         Top             =   1920
         Width           =   1005
      End
      Begin VB.TextBox txtBloodGroupeq 
         Alignment       =   2  'Center
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
         Left            =   2820
         TabIndex        =   18
         Text            =   "="
         Top             =   1500
         Width           =   1005
      End
      Begin VB.ComboBox cmbDesignation 
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
         Left            =   3990
         TabIndex        =   17
         Top             =   1920
         Width           =   2775
      End
      Begin VB.ComboBox cmbBloodGroup 
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
         ItemData        =   "frmSearchEmployee.frx":0000
         Left            =   3990
         List            =   "frmSearchEmployee.frx":001C
         TabIndex        =   16
         Top             =   1500
         Width           =   2775
      End
      Begin VB.CheckBox chkDesignation 
         Caption         =   "Designation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1110
         TabIndex        =   15
         Top             =   1950
         Width           =   1485
      End
      Begin VB.CheckBox chkBloodGroup 
         Caption         =   "BloodGroup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1110
         TabIndex        =   14
         Top             =   1545
         Width           =   1455
      End
      Begin VB.CheckBox chkDepartment 
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1110
         TabIndex        =   13
         Top             =   270
         Width           =   1545
      End
      Begin VB.CheckBox chkFirstName 
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1110
         TabIndex        =   12
         Top             =   735
         Width           =   1545
      End
      Begin VB.ComboBox cmbDepartment 
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
         Left            =   3990
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   240
         Width           =   2775
      End
      Begin VB.ComboBox cmbFirstNameOpr 
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
         ItemData        =   "frmSearchEmployee.frx":0042
         Left            =   2820
         List            =   "frmSearchEmployee.frx":004C
         TabIndex        =   10
         Top             =   660
         Width           =   1005
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "="
         Top             =   240
         Width           =   1005
      End
      Begin VB.TextBox txtFirstname 
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
         Left            =   3990
         TabIndex        =   8
         Top             =   660
         Width           =   2775
      End
      Begin VB.CheckBox chkLastName 
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1110
         TabIndex        =   7
         Top             =   1140
         Width           =   1545
      End
      Begin VB.ComboBox cmbLastNameOpr 
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
         ItemData        =   "frmSearchEmployee.frx":0059
         Left            =   2820
         List            =   "frmSearchEmployee.frx":0063
         TabIndex        =   6
         Top             =   1080
         Width           =   1005
      End
      Begin VB.TextBox txtLastname 
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
         Left            =   3990
         TabIndex        =   5
         Top             =   1080
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3345
      Left            =   270
      TabIndex        =   2
      Top             =   3540
      Width           =   7875
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdResult 
         Height          =   2805
         Left            =   120
         TabIndex        =   3
         Top             =   330
         Width           =   7635
         _ExtentX        =   13467
         _ExtentY        =   4948
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         AllowUserResizing=   3
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
         _Band(0).Cols   =   8
      End
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "S&earch"
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
      Left            =   3180
      TabIndex        =   1
      Top             =   3030
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "C&ancel"
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
      Left            =   4290
      TabIndex        =   0
      Top             =   3030
      Width           =   1005
   End
End
Attribute VB_Name = "frmSearchEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub chkBloodGroup_Click()
    If chkBloodGroup.Value = 0 Then cmbBloodGroup.Text = ""
    cmbBloodGroup.Enabled = chkBloodGroup.Value
End Sub

Private Sub chkDepartment_Click()
    If chkDepartment.Value = 0 Then cmbDepartment.Text = ""
    cmbDepartment.Enabled = chkDepartment.Value
End Sub

Private Sub chkDesignation_Click()
    If chkDesignation.Value = 0 Then cmbDesignation.Text = ""
    cmbDesignation.Enabled = chkDesignation.Value
End Sub

Private Sub chkFirstName_Click()
    If chkFirstName.Value = 0 Then
        cmbFirstNameOpr.Text = ""
        txtFirstname = ""
    End If
End Sub

Private Sub chkLastName_Click()
    If chkLastName.Value = 0 Then
        cmbLastNameOpr.Text = ""
        txtLastname.Text = ""
    End If
End Sub

Private Sub cmbBloodGroup_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmbDepartment_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmbDesignation_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmbFirstNameOpr_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmbLastNameOpr_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
    chkDepartment.Value = 0
    chkFirstName.Value = 0
    chkLastName.Value = 0
    chkBloodGroup.Value = 0
    chkDesignation.Value = 0
    cmbFirstNameOpr.Text = ""
    cmbLastNameOpr.Text = ""
    cmbDepartment.Text = ""
    txtFirstname.Text = ""
    txtLastname.Text = ""
    cmbBloodGroup.Text = ""
    cmbDesignation.Text = ""
    grdResult.Clear
    Frame2.Caption = " Result "
    Call mFormatGrid
    chkDepartment.SetFocus
    cmbDepartment.Enabled = False
    cmbDesignation.Enabled = False
    cmbBloodGroup.Enabled = False
End Sub

Private Sub cmdSearch_Click()
On Error GoTo SearchError
    
    Dim strCriteria As String, rsSearch As New ADODB.Recordset
    
'    If chkDepartment.Value = 0 And chkFirstName.Value = 0 And chkLastName.Value = 0 Then
'        MsgBox "Please specify criteria to search Employee", vbInformation, "Search"
'        Exit Sub
    If chkDepartment.Value <> 0 And Len(Trim(cmbDepartment.Text)) = 0 Then
        MsgBox "Please specify Department Name ", vbInformation, "Search"
        cmbDepartment.SetFocus
        Exit Sub
    ElseIf chkBloodGroup.Value <> 0 And Len(Trim(cmbBloodGroup.Text)) = 0 Then
        MsgBox "Please specify Bloodgroup ", vbInformation, "Search"
        cmbBloodGroup.SetFocus
        Exit Sub
    ElseIf chkDesignation.Value <> 0 And Len(Trim(cmbDesignation.Text)) = 0 Then
        MsgBox "Please specify Designation", vbInformation, "Search"
        cmbDesignation.SetFocus
        Exit Sub
    ElseIf chkFirstName.Value <> 0 Then
        If Len(Trim(txtFirstname)) = 0 Then
            MsgBox "Please specify Employee Name or starting alphabate", vbInformation, "Search"
            txtFirstname.SetFocus
            Exit Sub
        ElseIf Len(Trim(cmbFirstNameOpr.Text)) = 0 Then
            MsgBox "Please select '=' or Like operator", vbInformation, "Search"
            cmbFirstNameOpr.SetFocus
            Exit Sub
        End If
    ElseIf chkLastName.Value <> 0 Then
        If Len(Trim(txtLastname)) = 0 Then
            MsgBox "Please specify Employee Last Name or starting alphabate", vbInformation, "Search"
            txtLastname.SetFocus
            Exit Sub
        ElseIf Len(Trim(cmbLastNameOpr.Text)) = 0 Then
            MsgBox "Please select '=' or Like operator", vbInformation, "Search"
            cmbLastNameOpr.SetFocus
            Exit Sub
        End If
    End If
    
    If chkDepartment.Value <> 0 Then
        strCriteria = " where Department = '" & cmbDepartment & "'"
    ElseIf chkFirstName.Value <> 0 Then
        If Len(strCriteria) > 0 Then
            strCriteria = strCriteria & " and efname " & cmbFirstNameOpr.Text & "'" & txtFirstname.Text & IIf(cmbLastNameOpr.Text <> "=", "%", "") & "'"
        Else
            strCriteria = " where efname " & cmbFirstNameOpr.Text & "'" & txtFirstname.Text & IIf(cmbLastNameOpr.Text <> "=", "%", "") & "'"
        End If
    ElseIf chkLastName.Value <> 0 Then
        If Len(strCriteria) > 0 Then
            strCriteria = strCriteria & " and elname " & cmbLastNameOpr.Text & "'" & txtLastname.Text & IIf(cmbLastNameOpr.Text <> "=", "%", "") & "'"
        Else
            strCriteria = " where elname " & cmbLastNameOpr.Text & "'" & txtLastname.Text & IIf(cmbLastNameOpr.Text <> "=", "%", "") & "'"
        End If
    ElseIf chkBloodGroup <> 0 Then
        If Len(strCriteria) > 0 Then
            strCriteria = strCriteria & " and bloodgroup ='" & cmbBloodGroup.Text & "'"
        Else
            strCriteria = " where bloodgroup='" & cmbBloodGroup.Text & "'"
        End If
    ElseIf chkDesignation <> 0 Then
        If Len(strCriteria) > 0 Then
            strCriteria = strCriteria & " and designation='" & cmbDesignation.Text & "'"
        Else
            strCriteria = " where designation='" & cmbDesignation.Text & "'"
        End If
    End If
    
    
    
    rsSearch.Open "select * from searchemployee " & strCriteria, gStrConnectionString, adOpenKeyset, adLockOptimistic
    
    If rsSearch.RecordCount > 0 Then
        Set grdResult.Recordset = rsSearch
        Frame2.Caption = " Result: " & rsSearch.RecordCount & " Record(s) Found"
    Else
        Frame2.Caption = " Result:  0 Record(s) Found "
        grdResult.Clear
    End If
    Call mFormatGrid
Exit Sub
SearchError:
    MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, vbCritical, "Search Employee"
    Err.Clear
End Sub

Private Sub Form_Load()
On Error GoTo LoadError
    Dim rsDepartment As New ADODB.Recordset, i As Integer
        
    Call gFormCenter(Me)
        
    rsDepartment.Open "select * from m_department where status=1", gStrConnectionString, adOpenKeyset, adLockOptimistic
    cmbDepartment.Clear
    
    If rsDepartment.RecordCount > 0 Then
        For i = 0 To rsDepartment.RecordCount - 1
            cmbDepartment.AddItem rsDepartment.Fields(1)
            rsDepartment.MoveNext
        Next
    End If
    
    rsDepartment.Close
    rsDepartment.Open "select * from m_designation where status=1", gStrConnectionString, adOpenKeyset, adLockOptimistic
    cmbDesignation.Clear
    
    If rsDepartment.RecordCount > 0 Then
        For i = 0 To rsDepartment.RecordCount - 1
            cmbDesignation.AddItem rsDepartment.Fields(1)
            rsDepartment.MoveNext
        Next
    End If
    
Exit Sub
LoadError:
    MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, vbCritical, "Search Employee"
    Err.Clear
End Sub

Public Sub mFormatGrid()
    With grdResult
        .Row = 0
        .Col = 0
        .Text = "ID"
        .ColWidth(0) = 500
        .Col = 1
        .Text = "First Name"
        .ColWidth(1) = 1100
        .Col = 2
        .Text = "Middle Name"
        .ColWidth(2) = 1100
        .Col = 3
        .Text = "Last Name"
        .Col = 4
        .Text = "Department"
        .Col = 5
        .Text = "Designation"
        .Col = 6
        .Text = "Contact"
        .Col = 7
        .Text = "DOJ"
        .ColWidth(7) = 1050
    End With
End Sub

