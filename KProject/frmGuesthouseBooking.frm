VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmGuesthouseBooking 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Guesthouse Booking"
   ClientHeight    =   4020
   ClientLeft      =   2040
   ClientTop       =   2040
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Request Detail"
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
      Left            =   240
      TabIndex        =   9
      Top             =   270
      Width           =   6135
      Begin MSComCtl2.DTPicker dtpArrivalTime 
         Height          =   360
         Left            =   4410
         TabIndex        =   6
         Top             =   1800
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
         Format          =   19726338
         UpDown          =   -1  'True
         CurrentDate     =   38821
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   360
         Left            =   4410
         TabIndex        =   8
         Top             =   2220
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
         Format          =   19726337
         CurrentDate     =   38821
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   360
         Left            =   1350
         TabIndex        =   7
         Top             =   2220
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
         Format          =   19726337
         CurrentDate     =   38821
      End
      Begin VB.CommandButton cmdRoomList 
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
         Left            =   5190
         TabIndex        =   4
         Top             =   1380
         Width           =   735
      End
      Begin VB.TextBox txtRoom 
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
         Left            =   4410
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1365
         Width           =   705
      End
      Begin VB.CommandButton cmdGuesthouseList 
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
         Left            =   2430
         TabIndex        =   3
         Top             =   1350
         Width           =   735
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
         Left            =   5190
         TabIndex        =   2
         Top             =   930
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtpRequestDate 
         Height          =   360
         Left            =   4410
         TabIndex        =   1
         Top             =   480
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
         Format          =   19726337
         CurrentDate     =   38821
      End
      Begin VB.TextBox txtName 
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
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   900
         Width           =   3765
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
         Left            =   1350
         TabIndex        =   11
         Top             =   480
         Width           =   825
      End
      Begin VB.TextBox txtGuesthouse 
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
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1350
         Width           =   1005
      End
      Begin VB.ComboBox cmbGuestType 
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
         ItemData        =   "frmGuesthouseBooking.frx":0000
         Left            =   1350
         List            =   "frmGuesthouseBooking.frx":000A
         TabIndex        =   5
         Top             =   1785
         Width           =   1635
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Arrival Time:"
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
         Left            =   3210
         TabIndex        =   22
         Top             =   1822
         Width           =   1125
      End
      Begin VB.Label Label5 
         Caption         =   "To Date:"
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
         Left            =   3540
         TabIndex        =   21
         Top             =   2250
         Width           =   945
      End
      Begin VB.Label Label4 
         Caption         =   "Room:"
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
         Left            =   3720
         TabIndex        =   19
         Top             =   1380
         Width           =   615
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
         Left            =   3840
         TabIndex        =   18
         Top             =   510
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "From Date:"
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
         Left            =   240
         TabIndex        =   17
         Top             =   2280
         Width           =   990
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Guest Type:"
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
         Left            =   150
         TabIndex        =   16
         Top             =   1845
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Guesthouse:"
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
         Left            =   135
         TabIndex        =   15
         Top             =   1410
         Width           =   1125
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Emp Name:"
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
         Left            =   195
         TabIndex        =   14
         Top             =   975
         Width           =   1065
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
         Left            =   915
         TabIndex        =   13
         Top             =   540
         Width           =   360
      End
   End
   Begin KProject.ActionButton ActionButton1 
      Height          =   495
      Left            =   690
      TabIndex        =   0
      Top             =   3240
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
Attribute VB_Name = "frmGuesthouseBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub ActionButton1_AddClick(iNextNo As Integer)

    Call mClearControl
    txtId = iNextNo + 1
    Call mResetControl(True)
    dtpRequestDate.SetFocus
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
    frmSelect.strSQL = "select ghb_id as [Req ID] ,ghb_date as [Booking Dt],emp_fname as [By Emp] from view_guesthousebooking" & _
                        " where status=1"
    gBlGuesthouseBooking = True
    frmSelect.Show vbModal
    
    If gIntGuesthouseBooking > 0 Then
        Dim rsDisplay As New ADODB.Recordset
        
        rsDisplay.Open "select * from view_guesthousebooking where ghb_id=" & gIntGuesthouseBooking, _
                        gStrConnectionString, adOpenKeyset, adLockOptimistic
        txtId = rsDisplay.Fields(0)
        dtpRequestDate.Value = rsDisplay.Fields(1)
        txtName = rsDisplay.Fields("emp_fname")
        txtGuesthouse = rsDisplay("gh_location")
        txtRoom = rsDisplay.Fields(3)
        cmbGuestType.Text = rsDisplay.Fields(5)
        dtpArrivalTime.Value = rsDisplay.Fields("ghb_arrival_time")
        dtpFromDate.Value = rsDisplay.Fields(6)
        dtpToDate.Value = rsDisplay.Fields(7)
        gintListGuesthouse = rsDisplay.Fields("gh_id")
        gintListEmployee = rsDisplay.Fields("emp_id")
        
        
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
On Error GoTo UpdateError
    If Not ActionButton1.blAdd Then ActionButton1.blModify = True

    If Len(Trim(txtName)) = 0 Then
        MsgBox "Please select Employee", vbInformation, "Update"
        cmdEmployeeList.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtGuesthouse)) = 0 Then
        MsgBox "Please select Guesthouse ", vbInformation, "Update"
        cmdGuesthouseList.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtRoom)) = 0 Then
        MsgBox "Please select Room", vbInformation, "Update"
        cmdRoomList.SetFocus
        Exit Sub
    ElseIf Len(Trim(cmbGuestType.Text)) = 0 Then
        MsgBox "Please select Guest Type (Corporate or Personal)", vbInformation, "Update"
        cmbGuestType.SetFocus
        Exit Sub
    End If

    Dim rsCheck As New ADODB.Recordset, strSQL As String

    strSQL = "select * from t_guesthouse_Booking where gh_id=" & gintListGuesthouse & " and " & _
            " ghd_room_no ='" & txtRoom & "' and (ghb_duration_from between #" & dtpFromDate.Value & _
            "# and #" & dtpToDate.Value & "# or ghb_duration_to between #" & dtpFromDate.Value & _
            "# and #" & dtpToDate.Value & "# ) and ghb_id <>" & txtId
    Debug.Print strSQL
    rsCheck.Open strSQL, gStrConnectionString, adOpenKeyset, adLockOptimistic
    
    If rsCheck.RecordCount > 0 Then
        MsgBox "Room already booked by Other" & vbCrLf & "Booking is not updated", vbExclamation, "Update"
        ActionButton1.blModify = False
        ActionButton1.blSave = False
        ActionButton1.cmdCancel_Click
        Exit Sub
    Else
        ActionButton1.blSave = True
        'need to replace zero with guserid if guserid is admin
        ActionButton1.SaveSql = "Insert into t_guesthouse_booking values (" & txtId & _
                ",#" & dtpRequestDate & "#," & gintListGuesthouse & ",'" & txtRoom & _
                "'," & gintListEmployee & ",'" & cmbGuestType.Text & "',#" & dtpFromDate.Value & _
                "#,#" & dtpToDate.Value & "#,0,#" & dtpArrivalTime.Value & "#,1)"
        Call mResetControl(False)
    End If
Exit Sub
UpdateError:
    MsgBox "Error :" & Err.Number & vbCrLf & "Desc :" & Err.Description, vbCritical, "Update Error"
    Err.Clear
End Sub

Private Sub cmbGuestType_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdEmployeeList_Click()
    gBlListEmployee = True
    frmList.strSQL = "select emp_id as ID,emp_fname as [First Name],emp_lname as [Last Name],dept_name as Department from view_employee where status=1"
    frmList.Show vbModal
    txtName.Locked = False
    txtName = gstrListEmployee
    txtName.Locked = True
    gBlListEmployee = False
End Sub

Private Sub cmdGuesthouseList_Click()
    gBlListGuesthouse = True
    frmList.strSQL = "select gh_id as ID, gh_location as Guesthouse from m_guesthouse where status=1"
    frmList.Show vbModal
    txtGuesthouse.Locked = False
    txtGuesthouse = gstrListGuesthouse
    txtGuesthouse.Locked = True
    gBlListGuesthouse = False
End Sub

Private Sub cmdRoomList_Click()
    gBlListRoom = True
    frmList.strSQL = "select gh_id as [Ghouse ID], ghd_Room_no as Room from m_guesthouse_detail where status=1 and gh_id=" & gintListGuesthouse
    frmList.Show vbModal
    txtRoom.Locked = False
    txtRoom = gstrListRoom
    txtRoom.Locked = True
    gBlListRoom = False
End Sub

Private Sub Form_Load()

    Call gFormCenter(Me)
    With ActionButton1
        .ConnectionString = gStrConnectionString
        .PrimaryKeyField = "Ghb_id"
        .Table = "t_guesthouse_booking"
    End With
    Call mResetControl(False)
End Sub

Public Sub mResetControl(ByVal blEnable As Boolean)
    txtName.Enabled = blEnable
    dtpRequestDate.Enabled = blEnable
    txtGuesthouse.Enabled = blEnable
    txtRoom.Enabled = blEnable
    cmbGuestType.Enabled = blEnable
    dtpArrivalTime.Enabled = blEnable
    dtpFromDate.Enabled = blEnable
    dtpToDate.Enabled = blEnable
    cmdEmployeeList.Enabled = blEnable
    cmdGuesthouseList.Enabled = blEnable
    cmdRoomList.Enabled = blEnable
End Sub

Public Sub mClearControl()
    txtId.Text = ""
    txtName.Locked = False
    txtName.Text = ""
    txtName.Locked = True
    dtpRequestDate.Value = Date
    txtGuesthouse.Locked = False
    txtGuesthouse.Text = ""
    txtGuesthouse.Locked = True
    txtRoom.Locked = False
    txtRoom.Text = ""
    txtRoom.Locked = True
    cmbGuestType.Text = ""
    dtpArrivalTime.Value = "00:00:00"
    dtpFromDate.Value = Date
    dtpToDate.Value = Date
End Sub


