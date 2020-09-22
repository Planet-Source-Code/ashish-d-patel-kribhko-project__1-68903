VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List"
   ClientHeight    =   4095
   ClientLeft      =   2745
   ClientTop       =   2745
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
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
      Height          =   360
      Left            =   2250
      TabIndex        =   2
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "O&K"
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
      Left            =   1290
      TabIndex        =   1
      Top             =   3600
      Width           =   915
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2925
      Left            =   180
      TabIndex        =   0
      Top             =   510
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   5159
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
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Dim rs As New ADODB.Recordset
Public strSql As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
       
Dim strId As Integer
Dim strValue As String
       
    If MSHFlexGrid1.Row = 0 Then Exit Sub
       
    With MSHFlexGrid1
        If gBlListOrder Then
            .Col = 3
            gintListOrder = .Text
            .Col = 1
            gstrListOrder = .Text
        End If
        If gBlListOrder Then GoTo LocalExit
        .Col = 0
            strId = .Text
        .Col = 1
            strValue = .Text
        If gBlListSubDept Then
            .Col = 2
            strValue = .Text
        End If
    End With

    If gBlListDept Then
        gintListDeptId = strId
        gstrListDeptName = strValue
    ElseIf gBlListDesg Then
        gintListDesgId = strId
        gstrListDesgName = strValue
    ElseIf gBlListShift Then
        gintListShiftId = strId
        gstrListShiftName = strValue
    ElseIf gBlListSubDept Then
        gintListSubDeptId = strId
        gstrListSubDeptName = strValue
    ElseIf gBlListEmployee Then
        gintListEmployee = strId
        gstrListEmployee = strValue
    ElseIf gBlListGuesthouse Then
        gintListGuesthouse = strId
        gstrListGuesthouse = strValue
    ElseIf gBlListRoom Then
        gintListRoom = strId
        gstrListRoom = strValue
    ElseIf gBlListAvailableVehicle Then
        gintListAvailableVehicle = strId
        gstrListAvailableVehicle = strValue
    ElseIf gBlListVehicleNum Then
        gintListVehicleNum = strId
        gstrListVehicleNum = strValue
    ElseIf gBlListCanteenItemType Then
        gintListCanteenItemType = strId
        gstrListCanteenItemType = strValue
    ElseIf gBlListApproveBy Then
        gintListApproveBy = strId
        gstrListApproveBy = strValue
    End If
LocalExit:
    Unload Me
End Sub

Private Sub Form_Load()
    
    Call gFormCenter(Me)
    
    rs.Open strSql, gStrConnectionString, adOpenKeyset, adLockOptimistic
    
    With MSHFlexGrid1
        Set .Recordset = rs
        If gBlListDept Or gBlListDesg Or gBlListGuesthouse Or gBlListRoom Or _
         gBlListAvailableVehicle Or gBlListVehicleNum Or gBlListCanteenItemType Then
            .ColWidth(0) = 500
            .ColWidth(1) = 3400
        ElseIf gBlListSubDept Or gBlListShift Then
            .ColWidth(0) = 500
            .ColWidth(1) = 1650
            .ColWidth(2) = 1650
        ElseIf gBlListEmployee Or gBlListApproveBy Then
            .ColWidth(0) = 400
            .ColWidth(1) = 1000
            .ColWidth(2) = 1000
            .ColWidth(3) = 1200
        ElseIf gBlListOrder Then
            .ColWidth(0) = 1200
            .ColWidth(1) = 1100
            .ColWidth(2) = 1500
            .ColWidth(3) = 5
        End If
    End With
    
    rs.Close
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    gBlListDept = False
    gBlListDesg = False
    gBlListShift = False
    gBlListSubDept = False
    gBlListEmployee = False
    gBlListGuesthouse = False
    gBlListRoom = False
    gBlListAvailableVehicle = False
    gBlListOrder = False
    gBlListVehicleNum = False
    gBlListCanteenItemType = False
    gBlListApproveBy = False
End Sub

Private Sub MSHFlexGrid1_DblClick()
    If MSHFlexGrid1.Row <> 0 Then cmdOK_Click
End Sub
