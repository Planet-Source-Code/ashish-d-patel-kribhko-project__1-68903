VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmGuesthouse 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Guesthouse Master"
   ClientHeight    =   6075
   ClientLeft      =   2655
   ClientTop       =   1470
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Guesthouse Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   210
      TabIndex        =   6
      Top             =   240
      Width           =   5595
      Begin VB.ComboBox cmbAvailable 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmGuesthouse.frx":0000
         Left            =   3480
         List            =   "frmGuesthouse.frx":000A
         TabIndex        =   5
         Top             =   1980
         Width           =   1155
      End
      Begin VB.ComboBox cmbStatus 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmGuesthouse.frx":0017
         Left            =   2310
         List            =   "frmGuesthouse.frx":0021
         TabIndex        =   4
         Top             =   1980
         Width           =   1185
      End
      Begin VB.TextBox txtRoomNo 
         Appearance      =   0  'Flat
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
         Left            =   1200
         TabIndex        =   3
         Top             =   1980
         Width           =   1065
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdRoomDetail 
         Height          =   2565
         Left            =   1170
         TabIndex        =   11
         Top             =   2340
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   4524
         _Version        =   393216
         Cols            =   3
         FixedRows       =   0
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
         _Band(0).BandIndent=   1
         _Band(0).Cols   =   3
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
         _Band(0).ColHeader=   1
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
         Left            =   1170
         TabIndex        =   1
         Top             =   915
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
         Left            =   1170
         TabIndex        =   7
         Top             =   480
         Width           =   1245
      End
      Begin VB.TextBox txtRemark 
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
         Left            =   1170
         TabIndex        =   2
         Top             =   1350
         Width           =   4035
      End
      Begin VB.Label Label3 
         Caption         =   "Available Room(s) Details:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   180
         TabIndex        =   12
         Top             =   2310
         Width           =   855
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
         Index           =   1
         Left            =   315
         TabIndex        =   10
         Top             =   1380
         Width           =   765
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Location:"
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
         Left            =   270
         TabIndex        =   9
         Top             =   960
         Width           =   810
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
         TabIndex        =   8
         Top             =   540
         Width           =   240
      End
   End
   Begin KProject.ActionButton ActionButton1 
      Height          =   495
      Left            =   450
      TabIndex        =   0
      Top             =   5400
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
   Begin VB.Menu mnuRight 
      Caption         =   "Right"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove"
      End
   End
End
Attribute VB_Name = "frmGuesthouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim blGridClick As Boolean
Dim iRoomtop As Integer

Private Sub cmbAvailable_Click()
    If Len(cmbAvailable.Text) > 0 Then
        If Len(txtRoomNo.Text) > 0 And Len(cmbStatus.Text) > 0 Then
            With grdRoomDetail
                
                If blGridClick = False Then
                    Call mWriteToGrid(grdRoomDetail.Rows - 1)
                Else
                    Call mWriteToGrid(grdRoomDetail.Row)
                End If
                blGridClick = False
                Call mAddRowToGrid
            End With
        End If
    End If
End Sub

Private Sub cmbAvailable_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> 46 Then KeyCode = 0
End Sub

Private Sub cmbAvailable_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmbStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> 46 Then KeyCode = 0
End Sub

Private Sub cmbStatus_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Form_Load()
    
    Call gFormCenter(Me)
    
    With ActionButton1
        .ConnectionString = gStrConnectionString
        .PrimaryKeyField = "gh_id"
        .Table = "m_Guesthouse"
    End With
    
    iRoomtop = txtRoomNo.Top
    Call mResetControl(False)
    Call mInitialiseGrid
End Sub

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
    frmSelect.strSql = "select gh_id as ID ,gh_location as Location from m_guesthouse" & _
                        " where status=1"
    gBlGuesthouse = True
    frmSelect.Show vbModal
    
    If gIntGuestHOuse > 0 Then
        Dim rsDisplay As New ADODB.Recordset
        
        rsDisplay.Open "select gh_id,gh_location,gh_remark from m_guesthouse where gh_id=" & gIntGuestHOuse, _
                        gStrConnectionString, adOpenKeyset, adLockOptimistic
        txtId = rsDisplay.Fields(0)
        txtName = rsDisplay.Fields(1)
        txtRemark = rsDisplay.Fields(2)
       
        
        Dim rsDetail As New ADODB.Recordset
        
        rsDetail.Open "select ghd_room_no,ghd_room_status,ghd_room_allocated_flag from m_guesthouse_detail " & _
                        " where gh_id =" & txtId, gStrConnectionString, adOpenKeyset, adLockOptimistic
        
        Call mInitialiseGrid
        With grdRoomDetail
            Set .Recordset = rsDetail
            Call mReadFromGrid(1)
        End With
        
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

    If Len(Trim(txtName)) = 0 Then
        MsgBox "Guesthouse Location can not be left blank", vbInformation, "Update"
        Exit Sub
    End If

    ActionButton1.blSave = True
    ActionButton1.SaveSql = "Insert into m_guesthouse values (" & txtId & ",'" & txtName & "','" & txtRemark & "',1)"
    
    
    Dim cnDetail As New ADODB.Connection, strS As String, i As Integer
    Dim strRoomno As String, strStatus As String, strAllocated As String
    
    cnDetail.ConnectionString = gStrConnectionString
    cnDetail.Open
    
    If ActionButton1.blModify Then cnDetail.Execute "delete from m_guesthouse_detail where gh_id=" & txtId.Text
    
    With grdRoomDetail
        For i = 1 To .Rows - 1
            If Len(Trim(.TextMatrix(i, 0))) <> 0 Then
                strS = "Insert into m_guesthouse_detail values (" & txtId & ",'" & .TextMatrix(i, 0) & "','" & _
                        .TextMatrix(i, 1) & "','" & .TextMatrix(i, 2) & "',1)"
                Debug.Print strS
                cnDetail.Execute strS
            End If
        Next
        cnDetail.Close
    End With
    grdRoomDetail.SelectionMode = flexSelectionFree

    Call mResetControl(False)
    Exit Sub
UpdateError:
    MsgBox "Error: " & Err.Number & vbCrLf & "Desc: Error in Updation of record" & vbCrLf & Err.Description, vbCritical, "Update"
    Err.Clear
End Sub


Public Sub mResetControl(ByVal blEnable As Boolean)
    txtName.Enabled = blEnable
    txtRemark.Enabled = blEnable
    txtRoomNo.Enabled = blEnable
    cmbStatus.Enabled = blEnable
    cmbAvailable.Enabled = blEnable
    grdRoomDetail.Enabled = blEnable
End Sub

Public Sub mClearControl()
    txtId.Text = ""
    txtName.Text = ""
    txtRemark.Text = ""
    txtRoomNo.Text = ""
    cmbStatus.Text = ""
    cmbAvailable.Text = ""
    grdRoomDetail.Clear
    Call mInitialiseGrid
End Sub

Public Sub mInitialiseGrid()
 
 grdRoomDetail.Clear
    With grdRoomDetail
        .Rows = 2
        .BandDisplay = flexBandDisplayHorizontal
        .ColHeaderCaption(0, 0) = "Room No."
        .ColWidth(0) = 1300
        .ColHeaderCaption(0, 1) = "Status"
        .ColWidth(1) = 1300
        .ColHeaderCaption(0, 2) = "Available"
        .ColWidth(2) = 1300
        
        .Row = 1
        .Col = 0

        txtRoomNo.Left = .Left + .CellLeft
        txtRoomNo.Width = .CellWidth - 20
        txtRoomNo.Height = .CellHeight
        .Col = 1
        cmbStatus.Left = .Left + .CellLeft
        cmbStatus.Width = .CellWidth - 20
        .Col = 2
        cmbAvailable.Left = .Left + .CellLeft
        cmbAvailable.Width = .CellWidth - 20
    End With
End Sub

Private Sub grdRoomDetail_Click()
    blGridClick = True
    Call mReadFromGrid(grdRoomDetail.Row)

End Sub

Private Sub grdRoomDetail_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And ActionButton1.blModify Then
        PopupMenu mnuRight
    End If
End Sub

Private Sub mnuAdd_Click()
    blGridClick = False
    Call mAddRowToGrid
End Sub


Private Sub mnuRemove_Click()

    Dim iDel As Integer, i As Integer
    
    iDel = Val(InputBox("Enter Room No. to delete", "Delete", 0))
    If iDel > 0 Then
        grdRoomDetail.Col = 0
        For i = 1 To grdRoomDetail.Rows - 1
            grdRoomDetail.Row = i
            If iDel = Val(grdRoomDetail.Text) Then
                grdRoomDetail.RemoveItem (i - 1)
                grdRoomDetail.Refresh
                Exit For
            End If
        Next
    End If
End Sub

Public Sub mWriteToGrid(iRow As Integer)
    With grdRoomDetail
        .TextMatrix(iRow, 0) = txtRoomNo.Text
        .TextMatrix(iRow, 1) = cmbStatus.Text
        .TextMatrix(iRow, 2) = cmbAvailable.Text
    End With
End Sub

Public Sub mReadFromGrid(iRow As Integer)
    With grdRoomDetail
        txtRoomNo.Text = .TextMatrix(iRow, 0)
        cmbStatus.Text = .TextMatrix(iRow, 1)
        cmbAvailable.Text = .TextMatrix(iRow, 2)
        
        .Col = 0
        .Row = .Rows - 1
        If .Text = "" Then
            .Rows = .Rows - 1
        End If
        .Row = iRow
    End With
End Sub

Public Sub mAddRowToGrid()
    txtRoomNo.Text = ""
    cmbStatus.Text = ""
    cmbAvailable.Text = ""
   
    txtRoomNo.SetFocus
    
    If blGridClick Then Exit Sub
    
    grdRoomDetail.Rows = grdRoomDetail.Rows + 1
End Sub
