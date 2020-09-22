VERSION 5.00
Begin VB.UserControl ActionButton 
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5205
   LockControls    =   -1  'True
   ScaleHeight     =   480
   ScaleWidth      =   5205
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
      Height          =   405
      Left            =   4170
      TabIndex        =   4
      ToolTipText     =   "Cancel current Process"
      Top             =   30
      Width           =   1005
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "De&lete"
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
      Height          =   405
      Left            =   3135
      TabIndex        =   3
      ToolTipText     =   "Delete Record"
      Top             =   30
      Width           =   1005
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "&Modify"
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
      Left            =   2100
      TabIndex        =   2
      ToolTipText     =   "Modify Record"
      Top             =   30
      Width           =   1005
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
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
      Height          =   405
      Left            =   1065
      TabIndex        =   1
      ToolTipText     =   "Update/Save Record"
      Top             =   30
      Width           =   1005
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "A&dd"
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
      Left            =   30
      TabIndex        =   0
      ToolTipText     =   "Add New Record"
      Top             =   30
      Width           =   1005
   End
End
Attribute VB_Name = "ActionButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
Const m_def_ConnectionString = "0"
Const m_def_Table = "0"
Const m_def_PrimaryKeyField = "0"
Const m_def_SaveSql = "0" 'define Ashish 08/Apr/06

'Property Variables:
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
Dim m_ConnectionString As String
Dim m_Table As String
Dim m_PrimaryKeyField As String
Dim m_SaveSql As String

'Event Declarations:
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

'define Ashish 08/Apr/06
Event AddClick(iNextNo As Integer)  'MappingInfo=cmdAdd,cmdAdd,-1,Click
Event UpdateClick() 'MappingInfo=cmdUpdate,cmdUpdate,-1,Click
Attribute UpdateClick.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event ModifyClick() 'MappingInfo=cmdModify,cmdModify,-1,Click
Attribute ModifyClick.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DeleteClick() 'MappingInfo=cmdDelete,cmdDelete,-1,Click
Attribute DeleteClick.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event CancelClick() 'MappingInfo=cmdCancel,cmdCancel,-1,Click
Attribute CancelClick.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event AfterUpdateComplete()

'Variable define by Ashish Patel 08/Apr/06
Dim cnAction As New ADODB.Connection
Dim rsAction As New ADODB.Recordset
Public blSave As Boolean
Public blModify As Boolean
Public blAdd As Boolean
Public iModifyRecord As Integer
Public strDelete As String

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get BackColor() As Long
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get SaveSql() As String
        SaveSql = m_SaveSql
End Property

Public Property Let SaveSql(ByVal New_SaveSql As String)
    m_SaveSql = New_SaveSql
    PropertyChanged "SaveSql"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
     
End Sub

Private Sub cmdAdd_Click()
On Error GoTo LocalExit
    Dim iNext As Integer, strQ As String
    
    If rsAction.State = 1 Then rsAction.Close
    
    strQ = "select max(" & m_PrimaryKeyField & ") from " & m_Table
    rsAction.Open strQ, m_ConnectionString, adOpenKeyset, adLockOptimistic
    
    iNext = rsAction.Fields(0)
    
    RaiseEvent AddClick(iNext)
    rsAction.Close
    
    cmdUpdate.Enabled = True
    cmdModify.Enabled = False
    cmdDelete.Enabled = False
    cmdCancel.Enabled = True
    cmdAdd.Enabled = False
    blAdd = True
    Exit Sub
LocalExit:
    MsgBox "Error: " & Err.Number & vbCrLf & "Desc :" & Err.Description, vbCritical, "Active X Control Error"
    Err.Clear
End Sub
Private Sub cmdUPdate_Click()
On Error GoTo LocalExit
    
    RaiseEvent UpdateClick
    
    If Not blModify And Not blSave Then Exit Sub
    
    If cnAction.State = 1 Then cnAction.Close
    
    cnAction.ConnectionString = m_ConnectionString
    cnAction.Open
    If blModify Or blSave Then
        cnAction.BeginTrans
    
        If blModify Then cnAction.Execute "Delete from " & m_Table & " where " & m_PrimaryKeyField & " = " & iModifyRecord
        If blSave Then cnAction.Execute m_SaveSql
        cnAction.CommitTrans
    End If
        
    If Not blSave Then Exit Sub
    
    blModify = False
    blSave = False
    MsgBox "Record Updated ", vbInformation, "Update"

    cnAction.Close
    
    RaiseEvent AfterUpdateComplete
    blAdd = False
    cmdAdd.Enabled = True
    cmdModify.Enabled = True
    cmdUpdate.Enabled = False
    cmdDelete.Enabled = False
    cmdCancel.Enabled = True
    Exit Sub
LocalExit:
    cnAction.Cancel
    MsgBox "Error: " & Err.Number & vbCrLf & "Desc :" & Err.Description, vbCritical, "Active X Control Error"
    Err.Clear
End Sub

Private Sub cmdModify_Click()
    
    blModify = True
    blSave = True
    
    RaiseEvent ModifyClick
    
    cmdAdd.Enabled = False
    cmdUpdate.Enabled = True
    cmdModify.Enabled = False
    cmdDelete.Enabled = True
    cmdCancel.Enabled = True
    
    If blModify = False Then cmdCancel_Click
End Sub

Private Sub cmdDelete_Click()
On Error GoTo LocalExit
    
    RaiseEvent DeleteClick
    
    If MsgBox("Do you want to DELETE record " & strDelete, vbYesNo + vbQuestion, "Delete Confirm") = vbYes Then
        If cnAction.State = 1 Then cnAction.Close
            cnAction.ConnectionString = m_ConnectionString
            cnAction.Open
            cnAction.Execute "Update " & m_Table & " set status=0 where " & m_PrimaryKeyField & "=" & iModifyRecord
            cnAction.Close
            MsgBox "Record Deleted ", vbInformation, "Delete"
    End If
    
    cmdAdd.Enabled = True
    cmdModify.Enabled = True
    cmdUpdate.Enabled = False
    cmdDelete.Enabled = False
    cmdCancel.Enabled = True
    
    Exit Sub
LocalExit:
    MsgBox "Error: " & Err.Number & vbCrLf & "Desc :" & Err.Description, vbCritical, "Active X Control Error"
    Err.Clear
End Sub

Public Sub cmdCancel_Click()
    blModify = False
    blSave = False
    RaiseEvent CancelClick
    
    cmdAdd.Enabled = True
    cmdUpdate.Enabled = False
    cmdModify.Enabled = True
    cmdDelete.Enabled = False
    cmdCancel.Enabled = True
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get ConnectionString() As String
    ConnectionString = m_ConnectionString
End Property

Public Property Let ConnectionString(ByVal New_ConnectionString As String)
    m_ConnectionString = New_ConnectionString
    PropertyChanged "ConnectionString"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get Table() As String
Attribute Table.VB_Description = "select statement for table"
    Table = m_Table
End Property

Public Property Let Table(ByVal New_Table As String)
    m_Table = New_Table
    PropertyChanged "Table"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get PrimaryKeyField() As String
Attribute PrimaryKeyField.VB_Description = "primary key field"
    PrimaryKeyField = m_PrimaryKeyField
End Property

Public Property Let PrimaryKeyField(ByVal New_PrimaryKeyField As String)
    m_PrimaryKeyField = New_PrimaryKeyField
    PropertyChanged "PrimaryKeyField"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
    m_ConnectionString = m_def_ConnectionString
    m_Table = m_def_Table
    m_PrimaryKeyField = m_def_PrimaryKeyField
    m_SaveSql = m_def_SaveSql
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_ConnectionString = PropBag.ReadProperty("ConnectionString", m_def_ConnectionString)
    m_Table = PropBag.ReadProperty("Table", m_def_Table)
    m_PrimaryKeyField = PropBag.ReadProperty("PrimaryKeyField", m_def_PrimaryKeyField)
    m_SaveSql = PropBag.ReadProperty("SaveSql", m_def_SaveSql)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("ConnectionString", m_ConnectionString, m_def_ConnectionString)
    Call PropBag.WriteProperty("Table", m_Table, m_def_Table)
    Call PropBag.WriteProperty("PrimaryKeyField", m_PrimaryKeyField, m_def_PrimaryKeyField)
    Call PropBag.WriteProperty("SaveSql", m_SaveSql, m_def_SaveSql)
End Sub

