VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmVehicleUse 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vehicle Usage"
   ClientHeight    =   3885
   ClientLeft      =   2430
   ClientTop       =   2505
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Vehicle use Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   150
      TabIndex        =   7
      Top             =   210
      Width           =   6525
      Begin VB.TextBox txtToKm 
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
         Left            =   4830
         TabIndex        =   5
         Top             =   1830
         Width           =   1515
      End
      Begin VB.TextBox txtFromKm 
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
         Left            =   1770
         TabIndex        =   4
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmdVehicleList 
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
         Left            =   5610
         TabIndex        =   2
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtUsedBy 
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
         Height          =   375
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2280
         Width           =   3765
      End
      Begin VB.TextBox txtDriverName 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1770
         TabIndex        =   3
         Top             =   1350
         Width           =   4575
      End
      Begin VB.TextBox txtId 
         BackColor       =   &H80000018&
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
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   450
         Width           =   825
      End
      Begin VB.TextBox txtVehicleNo 
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
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   900
         Width           =   3765
      End
      Begin VB.CommandButton cmdEmployeeList 
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
         Left            =   5640
         TabIndex        =   6
         Top             =   2280
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtpUseDate 
         Height          =   360
         Left            =   4830
         TabIndex        =   1
         Top             =   450
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   19660801
         CurrentDate     =   38821
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "To (Km):"
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
         Left            =   3990
         TabIndex        =   16
         Top             =   1860
         Width           =   750
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "From (Km):"
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
         Left            =   690
         TabIndex        =   15
         Top             =   1830
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "No.:"
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
         Left            =   1290
         TabIndex        =   14
         Top             =   510
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vehicle No.:"
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
         Left            =   570
         TabIndex        =   13
         Top             =   945
         Width           =   1080
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Use By:"
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
         Left            =   945
         TabIndex        =   12
         Top             =   2280
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Date:"
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
         Left            =   4260
         TabIndex        =   11
         Top             =   540
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Driver Name:"
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
         TabIndex        =   10
         Top             =   1395
         Width           =   1185
      End
   End
   Begin KProject.ActionButton ActionButton1 
      Height          =   495
      Left            =   810
      TabIndex        =   0
      Top             =   3150
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
Attribute VB_Name = "frmVehicleUse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub ActionButton1_AddClick(iNextNo As Integer)

    Call mClearControl
    txtId = iNextNo + 1
    Call mResetControl(True)
    dtpUseDate.SetFocus
End Sub

Private Sub ActionButton1_CancelClick()
    Call mClearControl
    Call mResetControl(False)
End Sub

Private Sub ActionButton1_DeleteClick()
    With ActionButton1
        .iModifyRecord = txtId
        .strDelete = txtVehicleNo
    End With
    Call mClearControl
End Sub

Private Sub ActionButton1_ModifyClick()
    frmSelect.strSql = "select vhu_id as ID ,vhh_number_plate as [Vehicle],vhu_Date as [Date]," & _
                        " emp_fname as Employee  from view_vehicleuse where status=1"
    gBlVehicleUse = True
    frmSelect.Show vbModal
    
    If gIntVehicleUse > 0 Then
        Dim rsDisplay As New ADODB.Recordset
        
        rsDisplay.Open "select * from view_vehicleuse where vhu_id=" & gIntVehicleUse, _
                        gStrConnectionString, adOpenKeyset, adLockOptimistic
        txtId.Locked = False
        txtVehicleNo.Locked = False
        txtUsedBy.Locked = False
        txtId = rsDisplay.Fields(0)
        dtpUseDate.Value = rsDisplay.Fields(1)
        txtVehicleNo = rsDisplay.Fields("vhh_number_plate")
        txtFromKm.Text = rsDisplay.Fields(3)
        txtToKm.Text = rsDisplay.Fields(4)
        txtDriverName.Text = rsDisplay.Fields(5)
        txtUsedBy.Text = rsDisplay.Fields("emp_fname")
        gintListVehicleNum = rsDisplay.Fields("vhh_id")
        gintListEmployee = rsDisplay.Fields("Used_Emp_Id")
        txtId.Locked = True
        txtVehicleNo.Locked = True
        txtUsedBy.Locked = True
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
    
    If Len(Trim(txtVehicleNo)) = 0 Then
        MsgBox "Please select vehicle number", vbInformation, "Update"
        cmdVehicleList.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtDriverName)) = 0 Then
        MsgBox "Driver name can not be left blank", vbInformation, "Update"
        txtDriverName.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtFromKm)) = 0 Then
        MsgBox "Starting Km can not be left blank", vbInformation, "Update"
        txtFromKm.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtToKm)) = 0 Then
        MsgBox "End Km can not be left blank", vbInformation, "Update"
        txtToKm.SetFocus
        Exit Sub
    ElseIf Val(txtFromKm) >= Val(txtToKm) Then
        MsgBox "Wrong entry in StartKm and EndKm", vbInformation, "Update"
        txtFromKm.SetFocus
        Exit Sub
    ElseIf Len(txtUsedBy) = 0 Then
        MsgBox "Please select user of the vehicle", vbInformation, "Update"
        txtUsedBy.SetFocus
        Exit Sub
    End If

    ActionButton1.blSave = True
    ActionButton1.SaveSql = "Insert into t_vehicle_usage values (" & txtId & ",#" & _
                    dtpUseDate.Value & "#," & gintListVehicleNum & "," & txtFromKm & _
                    "," & txtToKm & ",'" & txtDriverName & "'," & gintListEmployee & ",1)"
    Call mResetControl(False)
End Sub

Private Sub cmdEmployeeList_Click()
    gBlListEmployee = True
    frmList.strSql = "select emp_id as ID,emp_fname as [First Name],emp_lname as [Last Name],dept_name as Department from view_employee where status=1"
    frmList.Show vbModal
    txtUsedBy.Locked = False
    txtUsedBy = gstrListEmployee
    txtUsedBy.Locked = True
    gBlListEmployee = False
End Sub

Private Sub cmdVehicleList_Click()

    gBlListVehicleNum = True
    frmList.strSql = "select vhh_id as ID,vhh_number_plate as [Vehicle Number Plate] from " & _
                        " view_vehiclenumberplate where status=1"
    frmList.Show vbModal
    txtVehicleNo.Locked = False
    txtVehicleNo = gstrListVehicleNum
    txtVehicleNo.Locked = True
    gBlListVehicleNum = False
End Sub

Private Sub Form_Load()

    Call gFormCenter(Me)
    With ActionButton1
        .ConnectionString = gStrConnectionString
        .PrimaryKeyField = "vhu_id"
        .Table = "t_vehicle_usage"
    End With
    Call mResetControl(False)
End Sub

Public Sub mResetControl(ByVal blEnable As Boolean)
    txtId.Enabled = blEnable
    txtVehicleNo.Enabled = blEnable
    txtDriverName.Enabled = blEnable
    txtFromKm.Enabled = blEnable
    txtToKm.Enabled = blEnable
    txtUsedBy.Enabled = blEnable
    cmdVehicleList.Enabled = blEnable
    cmdEmployeeList.Enabled = blEnable
End Sub

Public Sub mClearControl()
    txtId.Locked = False
    txtVehicleNo.Locked = False
    txtUsedBy.Locked = False
    txtId.Text = ""
    txtVehicleNo.Text = ""
    txtDriverName.Text = ""
    txtFromKm.Text = ""
    txtToKm.Text = ""
    txtUsedBy.Text = ""
    txtId.Locked = True
    txtVehicleNo.Locked = True
    txtUsedBy.Locked = True
    dtpUseDate.Value = Date
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    gIntEmployee = 0
    gIntVehicleUse = 0
End Sub
