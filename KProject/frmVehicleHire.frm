VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmVehicleHire 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vehicle Hire"
   ClientHeight    =   6945
   ClientLeft      =   2340
   ClientTop       =   810
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Vehicle Hiring and Owner Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   180
      TabIndex        =   13
      Top             =   150
      Width           =   6375
      Begin VB.TextBox txtNumberPlate 
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
         Left            =   1590
         TabIndex        =   2
         Top             =   1290
         Width           =   4575
      End
      Begin VB.ComboBox cmbAvailable 
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
         ItemData        =   "frmVehicleHire.frx":0000
         Left            =   4650
         List            =   "frmVehicleHire.frx":000A
         TabIndex        =   6
         Top             =   2130
         Width           =   1545
      End
      Begin VB.TextBox txtRCBookNo 
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
         Left            =   1590
         TabIndex        =   5
         Top             =   2130
         Width           =   1635
      End
      Begin VB.Frame Frame2 
         Caption         =   "Owner's Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   240
         TabIndex        =   20
         Top             =   2640
         Width           =   5955
         Begin VB.TextBox txtEmail 
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
            Left            =   1350
            TabIndex        =   12
            Top             =   2670
            Width           =   2445
         End
         Begin VB.TextBox txtOwnerMobile 
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
            Left            =   1350
            MaxLength       =   11
            TabIndex        =   11
            Top             =   2250
            Width           =   2445
         End
         Begin VB.TextBox txtOwnerhomepHone 
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
            Left            =   2190
            MaxLength       =   8
            TabIndex        =   10
            Top             =   1830
            Width           =   1605
         End
         Begin VB.TextBox txtOwnerSTD 
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
            Left            =   1350
            MaxLength       =   5
            TabIndex        =   9
            Top             =   1830
            Width           =   735
         End
         Begin VB.TextBox txtOwnerAdd 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   1350
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   750
            Width           =   4395
         End
         Begin VB.TextBox txtOwnerName 
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
            Left            =   1350
            TabIndex        =   7
            Top             =   330
            Width           =   4395
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "email:"
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
            TabIndex        =   25
            Top             =   2700
            Width           =   540
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Mobile:"
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
            Left            =   600
            TabIndex        =   24
            Top             =   2310
            Width           =   660
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Home Phone:"
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
            Left            =   30
            TabIndex        =   23
            Top             =   1890
            Width           =   1230
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Address:"
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
            Left            =   450
            TabIndex        =   22
            Top             =   900
            Width           =   810
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
            Index           =   3
            Left            =   660
            TabIndex        =   21
            Top             =   360
            Width           =   600
         End
      End
      Begin VB.TextBox txtVehicleName 
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
         Left            =   1590
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   870
         Width           =   3765
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
         Left            =   1590
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   450
         Width           =   825
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
         Left            =   5430
         TabIndex        =   1
         Top             =   870
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   360
         Left            =   4650
         TabIndex        =   4
         Top             =   1710
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
         Format          =   53739521
         CurrentDate     =   38821
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   360
         Left            =   1590
         TabIndex        =   3
         Top             =   1710
         Width           =   1635
         _ExtentX        =   2884
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
         Format          =   53739521
         CurrentDate     =   38821
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " Number Plate:"
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
         Left            =   195
         TabIndex        =   28
         Top             =   1320
         Width           =   1320
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "RC Book No:"
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
         Left            =   360
         TabIndex        =   27
         Top             =   2130
         Width           =   1170
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Available:"
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
         Left            =   3660
         TabIndex        =   26
         Top             =   2160
         Width           =   900
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
         Left            =   1170
         TabIndex        =   19
         Top             =   510
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "From:"
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
         Left            =   1020
         TabIndex        =   18
         Top             =   1770
         Width           =   510
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "To:"
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
         TabIndex        =   17
         Top             =   1740
         Width           =   300
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Vehicle Name:"
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
         Left            =   210
         TabIndex        =   16
         Top             =   915
         Width           =   1320
      End
   End
   Begin KProject.ActionButton ActionButton1 
      Height          =   495
      Left            =   750
      TabIndex        =   0
      Top             =   6300
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
Attribute VB_Name = "frmVehicleHire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub ActionButton1_AddClick(iNextNo As Integer)

    Call mClearControl
    txtId = iNextNo + 1
    Call mResetControl(True)
    cmdVehicleList.SetFocus
End Sub

Private Sub ActionButton1_CancelClick()
    Call mClearControl
    Call mResetControl(False)
End Sub

Private Sub ActionButton1_DeleteClick()
    With ActionButton1
        .iModifyRecord = txtId
        .strDelete = txtNumberPlate
    End With
    Call mClearControl
End Sub

Private Sub ActionButton1_ModifyClick()
    frmSelect.strSql = "select vhh_id as ID ,veh_make_comp as [Vehicle Name], vhh_owner_name as Owner from view_vehiclehire" & _
                        " where status=1"
    gBlVehicleHire = True
    frmSelect.Show vbModal
    
    If gIntVehicleHire > 0 Then
        Dim rsDisplay As New ADODB.Recordset
        
        rsDisplay.Open "select * from view_vehiclehire where vhh_id=" & gIntVehicleHire, _
                        gStrConnectionString, adOpenKeyset, adLockOptimistic
        txtId = rsDisplay.Fields(0)
        txtVehicleName = rsDisplay.Fields(1)
        dtpFromDate.Value = rsDisplay.Fields(2)
        dtpToDate.Value = rsDisplay.Fields(3)
        txtOwnerName = rsDisplay.Fields(4)
        txtOwnerAdd = rsDisplay.Fields(5)
        txtOwnerSTD = rsDisplay.Fields(6)
        txtOwnerhomepHone = rsDisplay.Fields(7)
        txtOwnerMobile = rsDisplay.Fields(8)
        txtEmail = rsDisplay.Fields(9)
        txtRCBookNo = rsDisplay.Fields(10)
        txtNumberPlate = rsDisplay.Fields("Vhh_number_plate")
        cmbAvailable.Text = rsDisplay.Fields(11)
        gintListAvailableVehicle = rsDisplay.Fields(13)
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
    
    If Len(Trim(txtVehicleName)) = 0 Then
        MsgBox "Please select Vehicle Name", vbInformation, "Update"
        cmdVehicleList.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtRCBookNo)) = 0 Then
        MsgBox "Please enter RCbookno", vbInformation, "Update"
        txtRCBookNo.SetFocus
        Exit Sub
    ElseIf Len(Trim(cmbAvailable.Text)) = 0 Then
        MsgBox "Please select available status", vbInformation, "Update"
        cmbAvailable.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtOwnerName)) = 0 Then
        MsgBox "Owner Name can not be left blank", vbInformation, "Update"
        txtOwnerName.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtOwnerAdd)) = 0 Then
        MsgBox "Owner Address can not be left blank", vbInformation, "Update"
        txtOwnerAdd.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtOwnerhomepHone)) = 0 Then
        MsgBox "Owner Home Phone no. is required", vbInformation, "Update"
        txtOwnerhomepHone.SetFocus
        Exit Sub
    End If

    ActionButton1.blSave = True
    ActionButton1.SaveSql = "Insert into m_vehicle_hired values (" & txtId & "," & _
            gintListAvailableVehicle & ",'" & txtNumberPlate & "',#" & dtpFromDate.Value & _
            "#,#" & dtpToDate.Value & "#,'" & txtOwnerName & "','" & txtOwnerAdd & "','" & _
            txtOwnerSTD & "','" & txtOwnerhomepHone & "','" & txtOwnerMobile & "','" & _
            txtEmail & "','" & txtRCBookNo & "','" & cmbAvailable.Text & " ',1)"
    Call mResetControl(False)
End Sub

Private Sub cmbAvailable_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub

Private Sub cmbAvailable_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmdVehicleList_Click()
    gBlListAvailableVehicle = True
    frmList.strSql = "select veh_id as ID,veh_make_comp as Vehicle from m_vehicle where status=1 "
    frmList.Show vbModal
    txtVehicleName.Locked = False
    txtVehicleName = gstrListAvailableVehicle
    txtVehicleName.Locked = True
    gBlListAvailableVehicle = False
End Sub

Private Sub Form_Load()

    Call gFormCenter(Me)
    With ActionButton1
        .ConnectionString = gStrConnectionString
        .PrimaryKeyField = "Vhh_id"
        .Table = "m_Vehicle_hired"
    End With
    Call mResetControl(False)
End Sub

Public Sub mResetControl(ByVal blEnable As Boolean)
    txtId.Enabled = blEnable
    txtVehicleName.Enabled = blEnable
    dtpFromDate.Enabled = blEnable
    dtpToDate.Enabled = blEnable
    txtRCBookNo.Enabled = blEnable
    cmbAvailable.Enabled = blEnable
    txtOwnerAdd.Enabled = blEnable
    txtOwnerhomepHone.Enabled = blEnable
    txtOwnerMobile.Enabled = blEnable
    txtOwnerName.Enabled = blEnable
    txtOwnerSTD.Enabled = blEnable
    txtEmail.Enabled = blEnable
    cmdVehicleList.Enabled = blEnable
    txtNumberPlate.Enabled = blEnable
End Sub

Public Sub mClearControl()
    txtId.Locked = False
    txtId.Text = ""
    txtId.Locked = True
    txtVehicleName.Locked = False
    txtVehicleName.Text = ""
    txtVehicleName.Locked = True
    dtpFromDate.Value = Date
    dtpToDate.Value = Date
    txtRCBookNo.Text = ""
    cmbAvailable.Text = ""
    txtOwnerAdd.Text = ""
    txtOwnerhomepHone.Text = ""
    txtOwnerMobile.Text = ""
    txtOwnerSTD.Text = ""
    txtOwnerName.Text = ""
    txtEmail.Text = ""
    txtNumberPlate.Text = "GJ-5-"
End Sub
