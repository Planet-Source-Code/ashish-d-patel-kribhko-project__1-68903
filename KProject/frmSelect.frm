VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Record"
   ClientHeight    =   4995
   ClientLeft      =   1980
   ClientTop       =   1545
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
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
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   645
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Ca&ncel"
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
      Left            =   2910
      TabIndex        =   3
      Top             =   4500
      Width           =   1005
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "O&k"
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
      Left            =   1860
      TabIndex        =   2
      Top             =   4500
      Width           =   1005
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3795
      Left            =   120
      TabIndex        =   1
      Top             =   570
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   6694
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
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public strSql As String
Dim rs As New ADODB.Recordset

Private Sub cmdCancel_Click()
    
    gBlDepartment = False
    gBlDesignation = False
    gBlSubDepartment = False
    gBlHoliday = False
    gBlShift = False
    gBlProduct = False
    gBlVehicle = False
    gBlGuesthouse = False
    gBlEmployee = False
    gBlGuesthouseBooking = False
    gBlIssuePass = False
    gBlVehicleBooking = False
    gBlVehicleHire = False
    gBlRequirement = False
    gBlPaymentOfBill = False
    gBlVehicleUse = False
    gBlLeave = False
    gBlCanteenItemType = False
    gBlCanteenItem = False
    
    gIntDepartment = 0
    gIntDesignation = 0
    gIntSubDepartment = 0
    gIntHoliday = 0
    gIntShift = 0
    gIntProduct = 0
    gIntVehicle = 0
    gIntGuestHOuse = 0
    gIntEmployee = 0
    gIntGuesthouseBooking = 0
    gIntIssuePass = 0
    gIntVehicleBooking = 0
    gIntVehicleHire = 0
    gIntRequirement = 0
    gIntPaymentOfBill = 0
    gIntVehicleUse = 0
    gIntLeave = 0
    gIntCanteenItemType = 0
    gIntCanteenItem = 0
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If gBlDepartment Then gIntDepartment = Text1(0)
    If gBlDesignation Then gIntDesignation = Text1(0)
    If gBlSubDepartment Then gIntSubDepartment = Text1(0)
    If gBlHoliday Then gIntHoliday = Text1(0)
    If gBlShift Then gIntShift = Text1(0)
    If gBlProduct Then gIntProduct = Text1(0)
    If gBlVehicle Then gIntVehicle = Text1(0)
    If gBlGuesthouse Then gIntGuestHOuse = Text1(0)
    If gBlEmployee Then gIntEmployee = Text1(0)
    If gBlGuesthouseBooking Then gIntGuesthouseBooking = Text1(0)
    If gBlIssuePass Then gIntIssuePass = Text1(0)
    If gBlVehicleBooking Then gIntVehicleBooking = Text1(0)
    If gBlVehicleHire Then gIntVehicleHire = Text1(0)
    If gBlRequirement Then gIntRequirement = Text1(0)
    If gBlPaymentOfBill Then gIntPaymentOfBill = Text1(0)
    If gBlVehicleUse Then gIntVehicleUse = Text1(0)
    If gBlLeave Then gIntLeave = Text1(0)
    If gBlCanteenItemType Then gIntCanteenItemType = Text1(0)
    If gBlCanteenItem Then gIntCanteenItem = Text1(0)
    Unload Me
End Sub

Private Sub Form_Activate()
Dim i As Integer
    
    rs.Open strSql, gStrConnectionString, adOpenKeyset, adLockOptimistic
    
    Set MSHFlexGrid1.Recordset = rs
    MSHFlexGrid1.Refresh
        
    For i = 1 To rs.Fields.Count
            Load Text1(i)
            Text1(i).Visible = True
            Text1(i).Text = ""
            Text1(i).Alignment = 0
            Text1(i).Enabled = False
    Next
    
    If gBlDepartment Or gBlDesignation Or gBlGuesthouse Or gBlCanteenItemType Then
        MSHFlexGrid1.ColWidth(0) = Text1(0).Width
        MSHFlexGrid1.ColWidth(1) = 5090
        Text1(1).Left = Text1(0).Left + Text1(0).Width
        Text1(1).Width = 5060
    ElseIf gBlSubDepartment Or gBlHoliday Or gBlShift Or gBlProduct Or gBlVehicle Or _
                gBlGuesthouseBooking Or gBlIssuePass Or gBlVehicleBooking Or gBlVehicleHire _
                Or gBlRequirement Or gBlLeave Or gBlCanteenItem Then
        MSHFlexGrid1.ColWidth(0) = Text1(0).Width
        MSHFlexGrid1.ColWidth(1) = 2545
        MSHFlexGrid1.ColWidth(2) = 2530
        Text1(1).Left = Text1(0).Left + Text1(0).Width
        Text1(1).Width = 2545
        Text1(2).Left = Text1(1).Left + Text1(1).Width
        Text1(2).Width = 2545
    ElseIf gBlEmployee Or gBlPaymentOfBill Or gBlVehicleUse Then
        MSHFlexGrid1.ColWidth(0) = Text1(0).Width
        MSHFlexGrid1.ColWidth(1) = 1690
        MSHFlexGrid1.ColWidth(2) = 1690
        MSHFlexGrid1.ColWidth(3) = 1690
        Text1(1).Left = Text1(0).Left + Text1(0).Width
        Text1(1).Width = 1690
        Text1(2).Left = Text1(1).Left + Text1(1).Width
        Text1(2).Width = 1690
        Text1(3).Left = Text1(2).Left + Text1(2).Width
        Text1(3).Width = 1690
    End If
    Text1(0).Enabled = False
    rs.Close
    Call MSHFlexGrid1_Click
End Sub

Private Sub Form_Load()
    Call gFormCenter(Me)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    gBlDepartment = False
    gBlDesignation = False
    gBlSubDepartment = False
    gBlHoliday = False
    gBlShift = False
    gBlProduct = False
    gBlVehicle = False
    gBlGuesthouse = False
    gBlEmployee = False
    gBlGuesthouseBooking = False
    gBlIssuePass = False
    gBlVehicleBooking = False
    gBlVehicleHire = False
    gBlRequirement = False
    gBlPaymentOfBill = False
    gBlVehicleUse = False
    gBlLeave = False
    gBlCanteenItemType = False
    gBlCanteenItem = False
End Sub

Private Sub MSHFlexGrid1_Click()

With MSHFlexGrid1
    If .Row <> 0 Then
            .Col = 0
            Text1(0) = .Text
            .Col = 1
            Text1(1) = .Text
            
            If gBlSubDepartment Or gBlHoliday Or gBlShift Or gBlProduct Or gBlVehicle Or _
                gBlGuesthouseBooking Or gBlIssuePass Or gBlVehicleBooking Or gBlVehicleHire _
                Or gBlRequirement Or gBlLeave Or gBlCanteenItem Then
                .Col = 2
                Text1(2) = .Text
            ElseIf gBlEmployee Or gBlPaymentOfBill Or gBlVehicleUse Then
                .Col = 2
                Text1(2) = .Text
                .Col = 3
                Text1(3) = .Text
            End If
            .Col = 0
    End If
End With
End Sub
'code left incomplete
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Dim strS As String
    
    If gBlDepartment And Len(Trim(Text1(1))) > 0 Then
        strS = strSql & " and dept_name like '" & Text1(1) & "%'"
    ElseIf gBlDesignation And Len(Trim(Text1(1))) > 0 Then
        strS = strSql & " and desg_name like '" & Text1(1) & "%'"
    ElseIf gBlSubDepartment And Len(Trim(Text1(1))) > 0 Or Len(Trim(Text1(1))) > 0 Then
        strS = strSql & " and m_department.dept_name like '" & Text1(1) & "%' and m_sub_department.sub_dept_name like '" & Text1(2) & "%'"
    End If
    
    If Len(strS) > 0 Then
        If rs.State = 1 Then rs.Close
        Debug.Print strS
        rs.Open strS, gStrConnectionString, adOpenKeyset, adLockOptimistic
        
        MSHFlexGrid1.Clear
        Set MSHFlexGrid1.Recordset = rs
        MSHFlexGrid1.Refresh
    End If
End Sub
