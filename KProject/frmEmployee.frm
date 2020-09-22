VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEmployee 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employee Master"
   ClientHeight    =   7110
   ClientLeft      =   615
   ClientTop       =   645
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   Begin KProject.ActionButton ActionButton1 
      Height          =   495
      Left            =   1380
      TabIndex        =   0
      Top             =   6480
      Width           =   5235
      _ExtentX        =   9234
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6105
      Left            =   150
      TabIndex        =   17
      Top             =   240
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   10769
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General and Address Details"
      TabPicture(0)   =   "frmEmployee.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label34"
      Tab(0).Control(1)=   "Label32"
      Tab(0).Control(2)=   "txtCode"
      Tab(0).Control(3)=   "txtId"
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(5)=   "Frame2"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Contact and Office Details"
      TabPicture(1)   =   "frmEmployee.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame4 
         Caption         =   "Office Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3105
         Left            =   150
         TabIndex        =   60
         Top             =   2580
         Width           =   7275
         Begin VB.CommandButton cmdShiftList 
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
            Left            =   3300
            TabIndex        =   28
            Top             =   1590
            Width           =   735
         End
         Begin VB.CommandButton cmdDesignationList 
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
            Left            =   3300
            TabIndex        =   27
            Top             =   1170
            Width           =   735
         End
         Begin VB.CommandButton cmdSubDepartmentList 
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
            Left            =   3300
            TabIndex        =   26
            Top             =   750
            Width           =   735
         End
         Begin VB.CommandButton cmdDepartmentList 
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
            Left            =   3300
            TabIndex        =   25
            Top             =   330
            Width           =   735
         End
         Begin VB.ComboBox cmbPF 
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
            ItemData        =   "frmEmployee.frx":0038
            Left            =   5700
            List            =   "frmEmployee.frx":0042
            TabIndex        =   33
            Top             =   1140
            Width           =   1245
         End
         Begin VB.TextBox txtOtherAllowance 
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
            Left            =   5700
            TabIndex        =   32
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtHRA 
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
            Left            =   5700
            TabIndex        =   31
            Top             =   300
            Width           =   1455
         End
         Begin VB.TextBox txtBasic 
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
            TabIndex        =   30
            Top             =   2400
            Width           =   1395
         End
         Begin VB.TextBox txtDepartment 
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
            TabIndex        =   64
            Top             =   330
            Width           =   1650
         End
         Begin VB.TextBox txtSubDepartment 
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
            TabIndex        =   63
            Top             =   750
            Width           =   1650
         End
         Begin VB.TextBox txtDesignation 
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
            TabIndex        =   62
            Top             =   1170
            Width           =   1650
         End
         Begin VB.TextBox txtShift 
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
            TabIndex        =   61
            Top             =   1590
            Width           =   1650
         End
         Begin MSComCtl2.DTPicker dtpJoinDate 
            Height          =   345
            Left            =   1590
            TabIndex        =   29
            Top             =   2010
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   609
            _Version        =   393216
            Format          =   59375617
            CurrentDate     =   38820
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "PF:"
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
            Left            =   5310
            TabIndex        =   73
            Top             =   1200
            Width           =   300
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Other Allowance:"
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
            Left            =   4110
            TabIndex        =   72
            Top             =   750
            Width           =   1500
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "HRA:"
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
            Left            =   5130
            TabIndex        =   71
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Basic:"
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
            Left            =   930
            TabIndex        =   70
            Top             =   2430
            Width           =   555
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Join Date:"
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
            TabIndex        =   69
            Top             =   2040
            Width           =   900
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Shift Name:"
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
            Left            =   480
            TabIndex        =   68
            Top             =   1620
            Width           =   1020
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Designation:"
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
            Left            =   390
            TabIndex        =   67
            Top             =   1200
            Width           =   1125
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Subdepartment:"
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
            Left            =   90
            TabIndex        =   66
            Top             =   780
            Width           =   1425
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Department:"
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
            Left            =   420
            TabIndex        =   65
            Top             =   390
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Contact and Other Detail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   150
         TabIndex        =   53
         Top             =   810
         Width           =   7275
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
            ItemData        =   "frmEmployee.frx":004F
            Left            =   5700
            List            =   "frmEmployee.frx":006B
            TabIndex        =   24
            Top             =   1290
            Width           =   885
         End
         Begin VB.ComboBox cmbGender 
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
            ItemData        =   "frmEmployee.frx":0091
            Left            =   5700
            List            =   "frmEmployee.frx":009B
            TabIndex        =   22
            Top             =   420
            Width           =   1035
         End
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
            Height          =   375
            Left            =   1440
            TabIndex        =   21
            Top             =   1260
            Width           =   2385
         End
         Begin VB.TextBox txtMobile 
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
            Left            =   1440
            TabIndex        =   20
            Top             =   840
            Width           =   2385
         End
         Begin VB.TextBox txtHomePhone 
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
            Left            =   2220
            MaxLength       =   8
            TabIndex        =   19
            Top             =   420
            Width           =   1605
         End
         Begin VB.TextBox txtStdCode 
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
            Left            =   1440
            MaxLength       =   5
            TabIndex        =   18
            Top             =   420
            Width           =   705
         End
         Begin MSComCtl2.DTPicker dtpDOB 
            Height          =   375
            Left            =   5700
            TabIndex        =   23
            Top             =   840
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   661
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
            Format          =   59375617
            CurrentDate     =   38820
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Blood Group:"
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
            Left            =   4440
            TabIndex        =   59
            Top             =   1320
            Width           =   1185
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "DOB:"
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
            Left            =   5130
            TabIndex        =   58
            Top             =   930
            Width           =   480
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Gender:"
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
            Left            =   4890
            TabIndex        =   57
            Top             =   480
            Width           =   720
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Email:"
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
            Left            =   810
            TabIndex        =   56
            Top             =   1320
            Width           =   555
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Mobile No.:"
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
            Left            =   330
            TabIndex        =   55
            Top             =   870
            Width           =   1020
         End
         Begin VB.Label Label14 
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
            Left            =   150
            TabIndex        =   54
            Top             =   450
            Width           =   1230
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "General"
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
         Left            =   -74820
         TabIndex        =   43
         Top             =   1140
         Width           =   7275
         Begin VB.CommandButton cmdCopyAddress 
            Caption         =   "Copy To Parmenent Address"
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
            Left            =   4275
            TabIndex        =   11
            Top             =   2160
            Width           =   2865
         End
         Begin VB.ComboBox cmbSalutation 
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
            ItemData        =   "frmEmployee.frx":00AD
            Left            =   1065
            List            =   "frmEmployee.frx":00BD
            TabIndex        =   2
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtFname 
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
            Left            =   2085
            TabIndex        =   3
            Top             =   480
            Width           =   1635
         End
         Begin VB.TextBox txtMname 
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
            Left            =   3795
            TabIndex        =   4
            Top             =   480
            Width           =   1635
         End
         Begin VB.TextBox txtLname 
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
            Left            =   5505
            TabIndex        =   5
            Top             =   480
            Width           =   1635
         End
         Begin VB.TextBox txtAdd1 
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
            Left            =   1065
            TabIndex        =   6
            Top             =   900
            Width           =   6075
         End
         Begin VB.TextBox txtAdd2 
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
            Left            =   1065
            TabIndex        =   7
            Top             =   1320
            Width           =   6075
         End
         Begin VB.TextBox txtPincode 
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
            Left            =   5505
            TabIndex        =   9
            Top             =   1740
            Width           =   1635
         End
         Begin VB.ComboBox cmbCity 
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
            Left            =   1065
            TabIndex        =   8
            Top             =   1740
            Width           =   1995
         End
         Begin VB.ComboBox cmbState 
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
            Left            =   1065
            TabIndex        =   10
            Top             =   2160
            Width           =   1995
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Pincode:"
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
            Left            =   4635
            TabIndex        =   52
            Top             =   1800
            Width           =   795
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "State:"
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
            Left            =   495
            TabIndex        =   51
            Top             =   2220
            Width           =   510
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "City:"
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
            Left            =   630
            TabIndex        =   50
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Address2:"
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
            Left            =   90
            TabIndex        =   49
            Top             =   1380
            Width           =   915
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Address1:"
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
            Left            =   90
            TabIndex        =   48
            Top             =   960
            Width           =   915
         End
         Begin VB.Label Label28 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "SirName"
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
            Left            =   5535
            TabIndex        =   47
            Top             =   210
            Width           =   1605
         End
         Begin VB.Label Label29 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Middle Name"
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
            Left            =   3795
            TabIndex        =   46
            Top             =   210
            Width           =   1635
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label30 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   2085
            TabIndex        =   45
            Top             =   210
            Width           =   1635
         End
         Begin VB.Label Label31 
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
            Left            =   405
            TabIndex        =   44
            Top             =   540
            Width           =   600
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Permenant Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Left            =   -74820
         TabIndex        =   37
         Top             =   3720
         Width           =   7275
         Begin VB.ComboBox cmbPermState 
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
            Left            =   1050
            TabIndex        =   16
            Top             =   1560
            Width           =   1965
         End
         Begin VB.TextBox txtPermPincode 
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
            Left            =   5490
            TabIndex        =   15
            Top             =   1140
            Width           =   1635
         End
         Begin VB.ComboBox cmbPermCity 
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
            Left            =   1050
            TabIndex        =   14
            Top             =   1140
            Width           =   1965
         End
         Begin VB.TextBox txtPermAdd2 
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
            Left            =   1050
            TabIndex        =   13
            Top             =   720
            Width           =   6075
         End
         Begin VB.TextBox txtPermAdd1 
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
            Left            =   1050
            TabIndex        =   12
            Top             =   300
            Width           =   6075
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "State:"
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
            Left            =   480
            TabIndex        =   42
            Top             =   1620
            Width           =   510
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Pincode:"
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
            Left            =   4620
            TabIndex        =   41
            Top             =   1200
            Width           =   795
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "City:"
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
            TabIndex        =   40
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Address2:"
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
            Left            =   60
            TabIndex        =   39
            Top             =   780
            Width           =   915
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Address1:"
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
            Left            =   60
            TabIndex        =   38
            Top             =   360
            Width           =   915
         End
      End
      Begin VB.TextBox txtId 
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
         Left            =   -73740
         TabIndex        =   34
         Top             =   570
         Width           =   1185
      End
      Begin VB.TextBox txtCode 
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
         Left            =   -68910
         TabIndex        =   1
         Top             =   660
         Width           =   1185
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Code:"
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
         Left            =   -69480
         TabIndex        =   36
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label34 
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
         Left            =   -74010
         TabIndex        =   35
         Top             =   630
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'only number entry
'only alphabate entry

Option Explicit

Private Sub ActionButton1_AddClick(iNextNo As Integer)

    Call mClearControl
    txtId.Locked = False
    txtId = iNextNo + 1
    txtId.Locked = True
    Call mResetControl(True)
    Call mFillCombo
    txtCode.SetFocus
End Sub

Private Sub ActionButton1_AfterUpdateComplete()
    gBlListDept = False
    gBlListDesg = False
    gBlListShift = False
    gBlListSubDept = False
    gstrListDeptName = ""
    gstrListDesgName = ""
    gstrListShiftName = ""
    gstrListSubDeptName = ""
    gintListDeptId = 0
    gintListDesgId = 0
    gintListShiftId = 0
    gintListSubDeptId = 0
End Sub

Private Sub ActionButton1_CancelClick()
    Call mClearControl
    Call mResetControl(False)
End Sub

Private Sub ActionButton1_DeleteClick()
    With ActionButton1
        .iModifyRecord = txtId
        .strDelete = txtFname
    End With
    Call mClearControl
End Sub

Private Sub ActionButton1_ModifyClick()
    frmSelect.strSql = "select emp_id as ID ,emp_fname as Employee,dept_name as Department,desg_name as Designation from " & _
                        " m_employee me,m_department md,m_designation ms " & _
                        " where me.dept_id=md.dept_id and me.desg_id=ms.desg_id and me.status=1"
    gBlEmployee = True
    frmSelect.Show vbModal
    
    If gIntEmployee > 0 Then
        Dim rsDisplay As New ADODB.Recordset
        
        With rsDisplay
            .Open "select * from view_employee where emp_id=" & gIntEmployee, _
                        gStrConnectionString, adOpenKeyset, adLockOptimistic
        
            txtId = .Fields(0)
            txtCode = .Fields(1)
            cmbSalutation.Text = .Fields(2)
            txtFname = .Fields(3)
            txtMname = .Fields(4)
            txtLname = .Fields(5)
            txtAdd1 = .Fields(6)
            txtAdd2 = .Fields(7)
            cmbCity.Text = .Fields(8)
            cmbState.Text = .Fields(9)
            txtPincode = .Fields(10)
            txtPermAdd1 = .Fields(11)
            txtPermAdd2 = .Fields(12)
            cmbPermCity = .Fields(13)
            cmbPermState = .Fields(14)
            txtPermPincode = .Fields(15)
            cmbGender.Text = .Fields(16)
            dtpDOB.Value = .Fields(17)
            cmbBloodGroup.Text = .Fields(18)
            txtStdCode = .Fields(19)
            txtHomePhone = .Fields(20)
            txtMobile = .Fields(21)
            txtEmail = .Fields(22)
            txtDepartment = .Fields(23)
            txtSubDepartment = .Fields(24)
            txtDesignation = .Fields(25)
            txtShift = .Fields(26)
            dtpJoinDate.Value = .Fields(27)
            txtBasic = .Fields(29)
            txtHRA = .Fields(30)
            txtOtherAllowance = .Fields(31)
            cmbPF.Text = .Fields(32)
            gintListDeptId = .Fields("Dept_id")
            gintListDesgId = .Fields("desg_id")
            gintListSubDeptId = .Fields("sub_dept_id")
            gintListShiftId = .Fields("shf_id")
        End With
        ActionButton1.blModify = True
        ActionButton1.iModifyRecord = txtId
        Call mResetControl(True)
        txtId.Locked = True
    Else
        ActionButton1.blModify = False
        ActionButton1.blSave = False
        Call ActionButton1_CancelClick
    End If
End Sub

Private Sub ActionButton1_UpdateClick()
On Error GoTo UpdateError
        SSTab1.Tab = 0
    If Len(Trim(txtCode)) = 0 Then
        MsgBox "Employee Code can not be left blank", vbInformation, "Update"
        txtCode.SetFocus
        Exit Sub
    ElseIf Len(cmbSalutation.Text) = 0 Then
        MsgBox "Salution can not be left blank", vbInformation, "Update"
        cmbSalutation.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtFname)) = 0 Then
        MsgBox "First Name can not be left blank", vbInformation, "Update"
        txtFname.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtLname)) = 0 Then
        MsgBox "Last Name can not be left blank", vbInformation, "Update"
        txtLname.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtAdd1)) = 0 Then
        MsgBox "Address1 can not be left blank", vbInformation, "Update"
        txtAdd1.SetFocus
        Exit Sub
    ElseIf Len(cmbCity.Text) = 0 Then
        MsgBox "City Name can not be left blank", vbInformation, "Update"
        cmbCity.SetFocus
        Exit Sub
    ElseIf Len(cmbState.Text) = 0 Then
        MsgBox "State Name can not be left blank", vbInformation, "Update"
        cmbState.SetFocus
        Exit Sub
'    ElseIf Len(Trim(txtPincode)) = 0 Then
'        MsgBox "Pincode can not be left blank", vbInformation, "Update"
'        Exit Sub
    ElseIf Len(Trim(txtPermAdd1)) = 0 Then
        MsgBox "Permanent Address1 can not be left blank", vbInformation, "Update"
        txtPermAdd1.SetFocus
        Exit Sub
    ElseIf Len(Trim(cmbPermCity.Text)) = 0 Then
        MsgBox "Permanent Address City Name can not be left blank", vbInformation, "Update"
        cmbPermCity.SetFocus
        Exit Sub
    ElseIf Len(cmbPermState.Text) = 0 Then
        MsgBox "Permanent Address State Name can not be left blank", vbInformation, "Update"
        cmbPermState.SetFocus
        Exit Sub
    End If
    'elseif len(txtpermpincode
    SSTab1.Tab = 1
    If Len(cmbGender.Text) = 0 Then
        MsgBox "Please select Gender ( Male or Female )", vbInformation, "Update"
        cmbGender.SetFocus
        Exit Sub
    ElseIf Len(cmbBloodGroup.Text) = 0 Then
        MsgBox "Please select Blood Group", vbInformation, "Update"
        cmbBloodGroup.SetFocus
        Exit Sub
    'ElseIf Len(Trim(txtHomePhone)) = 0 Then
     '   MsgBox ""
    ElseIf Len(Trim(txtEmail)) > 0 And mValidEmail = False Then
        MsgBox "Please enter valid email address" & vbCrLf & " example c_ashish2000@yahoo.com", vbInformation, "Update"
        txtEmail.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtDepartment)) = 0 Then
        MsgBox "Please Select Department", vbInformation, "Update"
        cmdDepartmentList.SetFocus
        Exit Sub
    'remove those subdepartment from list, which are not related to department
    'ElseIf Len(Trim(txtSubDepartment)) > 0 And Len(Trim(txtDepartment)) = 0 Then
     '   MsgBox ""
    ElseIf Len(Trim(txtDesignation)) = 0 Then
        MsgBox "Please select Designation", vbInformation, "Update"
        cmdDesignationList.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtShift)) = 0 Then
        MsgBox "Please select Shift", vbInformation, "Update"
        cmdShiftList.SetFocus
        Exit Sub
    ElseIf Len(Trim(txtBasic)) = 0 Then
        MsgBox "Basic salary can not be left blank", vbInformation, "Update"
        txtBasic.SetFocus
        Exit Sub
    ElseIf Len(cmbPF.Text) = 0 Then
        MsgBox "Please select PF is applicable or not ?", vbInformation, "Update"
        cmbPF.SetFocus
        Exit Sub
    End If

    Dim rsCheck As New ADODB.Recordset
    
    rsCheck.Open "select * from m_employee where emp_code='" & Trim(txtCode) & "'", gStrConnectionString _
                , adOpenKeyset, adLockOptimistic
    If rsCheck.RecordCount > 0 Then
        MsgBox "Employee code already exist" & vbCrLf & "Please enter another employee code", vbInformation, "Update"
        SSTab1.Tab = 0
        txtCode.Text = ""
        txtCode.SetFocus
        Exit Sub
    End If

    ActionButton1.blSave = True
    ActionButton1.SaveSql = "Insert into m_Employee values (" & txtId & ",'" & txtCode & _
        "','" & cmbSalutation.Text & "','" & txtFname & "','" & txtMname & "','" & _
        txtLname & "','" & txtAdd1 & "','" & txtAdd2 & "','" & cmbCity.Text & "','" & _
        cmbState.Text & "','" & txtPincode & "','" & txtPermAdd1 & "','" & txtPermAdd2 & _
        "','" & cmbPermCity.Text & "','" & cmbPermState.Text & "','" & txtPermPincode & _
        "','" & cmbGender.Text & "',#" & dtpDOB.Value & "#,'" & cmbBloodGroup.Text & _
        "','" & txtStdCode & "','" & txtHomePhone & "','" & txtMobile & "','" & txtEmail & _
        "'," & gintListDeptId & "," & gintListSubDeptId & "," & gintListDesgId & _
        "," & gintListShiftId & ",#" & dtpJoinDate.Value & "#,'" & Left(txtFname, 1) & _
        "'," & Val(txtBasic) & "," & Val(txtHRA) & "," & Val(txtOtherAllowance) & ",'" & cmbPF.Text & "',1)"
    Call mResetControl(False)
    Exit Sub
UpdateError:
    MsgBox "Error: " & Err.Number & vbCrLf & "Desc :" & Err.Description, vbCritical, "Update Error"
    Err.Clear
End Sub

Private Sub cmbBloodGroup_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmbGender_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmbPF_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmbSalutation_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub

Private Sub cmdCopyAddress_Click()
    txtPermAdd1 = txtAdd1
    txtPermAdd2 = txtAdd2
    cmbPermCity.Text = cmbCity.Text
    cmbPermState.Text = cmbState.Text
    txtPermPincode = txtPincode
End Sub

Private Sub cmdDepartmentList_Click()
    gBlListDept = True
    frmList.strSql = "select dept_id as ID,dept_name as Department from m_department where status=1"
    frmList.Show vbModal
    txtDepartment.Locked = False
    txtDepartment = gstrListDeptName
    txtDepartment.Locked = True
    gBlListDept = False
End Sub

Private Sub cmdDesignationList_Click()
    gBlListDesg = True
    frmList.strSql = "select desg_id as ID,desg_name as Designation from m_designation where status=1"
    frmList.Show vbModal
    txtDesignation.Locked = False
    txtDesignation = gstrListDesgName
    txtDesignation.Locked = True
    gBlListDesg = False
End Sub

Private Sub cmdShiftList_Click()
    gBlListShift = True
    frmList.strSql = "select shf_id as ID,shf_name as Shift,shf_start_time as Start from m_shift where status=1"
    frmList.Show vbModal
    txtShift.Locked = False
    txtShift = gstrListShiftName
    txtShift.Locked = True
    gBlListShift = False
End Sub

Private Sub cmdSubDepartmentList_Click()
    gBlListSubDept = True
    frmList.strSql = "select sub_dept_id as ID,dept_name as Department,sub_dept_name as SubDepartment " & _
                    " from m_department md, m_sub_department ms where ms.status=1" & _
                    " and md.dept_id=ms.dept_id and ms.dept_id=" & gintListDeptId
    frmList.Show vbModal
    txtSubDepartment.Locked = False
    txtSubDepartment = gstrListSubDeptName
    txtSubDepartment.Locked = True
    gBlListSubDept = False
End Sub

Private Sub Form_Load()

    Call gFormCenter(Me)
    With ActionButton1
        .ConnectionString = gStrConnectionString
        .PrimaryKeyField = "emp_id"
        .Table = "m_Employee"
    End With
    Call mResetControl(False)
    SSTab1.Tab = 0
End Sub

Public Sub mResetControl(ByVal blEnable As Boolean)
    SSTab1.Tab = 0
    txtCode.Enabled = blEnable
    cmbSalutation.Enabled = blEnable
    txtFname.Enabled = blEnable
    txtMname.Enabled = blEnable
    txtLname.Enabled = blEnable
    txtAdd1.Enabled = blEnable
    txtAdd2.Enabled = blEnable
    cmbCity.Enabled = blEnable
    cmbState.Enabled = blEnable
    txtPincode.Enabled = blEnable
    txtPermAdd1.Enabled = blEnable
    txtPermAdd2.Enabled = blEnable
    cmbPermCity.Enabled = blEnable
    cmbPermState.Enabled = blEnable
    txtPermPincode.Enabled = blEnable
    txtStdCode.Enabled = blEnable
    txtHomePhone.Enabled = blEnable
    txtMobile.Enabled = blEnable
    txtEmail.Enabled = blEnable
    cmbGender.Enabled = blEnable
    cmbBloodGroup.Enabled = blEnable
    dtpDOB.Enabled = blEnable
    dtpJoinDate.Enabled = blEnable
    txtBasic.Enabled = blEnable
    txtHRA.Enabled = blEnable
    txtOtherAllowance.Enabled = blEnable
    cmbPF.Enabled = blEnable
    cmdDepartmentList.Enabled = blEnable
    cmdShiftList.Enabled = blEnable
    cmdSubDepartmentList.Enabled = blEnable
    cmdDesignationList.Enabled = blEnable
End Sub

Public Sub mClearControl()
    txtId.Locked = False
    txtId.Text = ""
    txtId.Locked = True
    txtCode.Text = ""
    cmbSalutation.Text = ""
    txtFname.Text = ""
    txtMname.Text = ""
    txtLname.Text = ""
    txtAdd1.Text = ""
    txtAdd2.Text = ""
    cmbCity.Text = ""
    cmbState.Text = ""
    txtPincode.Text = ""
    txtPermAdd1.Text = ""
    txtPermAdd2.Text = ""
    cmbPermCity.Text = ""
    cmbPermState.Text = ""
    txtPermPincode.Text = ""
    txtStdCode.Text = ""
    txtHomePhone.Text = ""
    txtMobile.Text = ""
    txtEmail.Text = ""
    cmbGender.Text = ""
    cmbBloodGroup.Text = ""
    dtpDOB.Value = DateAdd("yyyy", -21, Date)
    txtDepartment.Locked = False
    txtDesignation.Locked = False
    txtShift.Locked = False
    txtSubDepartment.Locked = False
    txtDepartment.Text = ""
    txtSubDepartment.Text = ""
    txtDesignation.Text = ""
    txtShift.Text = ""
    txtDepartment.Locked = True
    txtDesignation.Locked = True
    txtShift.Locked = True
    txtSubDepartment.Locked = True
    dtpJoinDate.Value = Date
    txtBasic.Text = ""
    txtHRA.Text = ""
    txtOtherAllowance.Text = ""
    cmbPF.Text = ""
End Sub

Public Sub mFillCombo()
On Error GoTo ComboFillError
    Dim i As Integer
    Dim rsFill As New ADODB.Recordset
        
        rsFill.Open "select distinct emp_city from m_employee", gStrConnectionString, adOpenKeyset, adLockOptimistic
        
        If rsFill.RecordCount > 0 Then
            cmbCity.Clear
            For i = 0 To rsFill.RecordCount - 1
                cmbCity.AddItem rsFill.Fields(0)
                rsFill.MoveNext
            Next
        End If
        
        rsFill.Close
        rsFill.Open "Select distinct emp_state from m_employee", gStrConnectionString, adOpenKeyset, adLockOptimistic
        
        If rsFill.RecordCount > 0 Then
            cmbState.Clear
            For i = 0 To rsFill.RecordCount - 1
                cmbState.AddItem rsFill.Fields(0)
                rsFill.MoveNext
            Next
        End If
        
        rsFill.Close
        rsFill.Open "select distinct emp_perm_city from m_employee", gStrConnectionString, adOpenKeyset, adLockOptimistic
        
        If rsFill.RecordCount > 0 Then
            cmbPermCity.Clear
            For i = 0 To rsFill.RecordCount - 1
                cmbPermCity.AddItem rsFill.Fields(0)
                rsFill.MoveNext
            Next
        End If
        
        rsFill.Close
        rsFill.Open "select distinct emp_perm_state from m_employee", gStrConnectionString, adOpenKeyset, adLockOptimistic
        
        If rsFill.RecordCount > 0 Then
            cmbPermState.Clear
            For i = 0 To rsFill.RecordCount - 1
                cmbPermState.AddItem rsFill.Fields(0)
                rsFill.MoveNext
            Next
        End If
        rsFill.Close
    Exit Sub
ComboFillError:
    MsgBox "Error :" & Err.Number & vbCrLf & Err.Description, vbCritical, "Combo Fill error"
    Err.Clear
End Sub


Public Function mValidEmail() As Boolean

Dim i As Integer, j As Integer

    i = InStr(1, txtEmail, "@", vbTextCompare)
    If i > 0 Then
        j = InStr(i + 1, txtEmail.Text, "@", vbTextCompare)
    End If
    
    If i = 0 Or j > 0 Then
        mValidEmail = False
    ElseIf i > 0 And j = 0 Then
        mValidEmail = True
    End If
End Function

Private Sub txtBasic_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyDelete Then
        KeyCode = 0
    End If

End Sub

Private Sub txtBasic_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyBack Then
        If KeyAscii < 47 Or KeyAscii > 58 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtHomePhone_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyDelete Then
        KeyCode = 0
    End If
End Sub

Private Sub txtHomePhone_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyBack Then
        If KeyAscii < 47 Or KeyAscii > 58 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtHRA_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyDelete Then
        KeyCode = 0
    End If

End Sub

Private Sub txtHRA_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyBack Then
        If KeyAscii < 47 Or KeyAscii > 58 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtMobile_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyDelete Then
        KeyCode = 0
    End If
End Sub

Private Sub txtMobile_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyBack Then
        If KeyAscii < 47 Or KeyAscii > 58 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtPermPincode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyDelete Then
        KeyCode = 0
    End If

End Sub

Private Sub txtPermPincode_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyBack Then
        If KeyAscii < 47 Or KeyAscii > 58 Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub txtPincode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyDelete Then
        KeyCode = 0
    End If
End Sub

Private Sub txtPincode_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyBack Then
        If KeyAscii < 47 Or KeyAscii > 58 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtStdCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyDelete Then
        KeyCode = 0
    End If
End Sub

Private Sub txtStdCode_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyBack Then
        If KeyAscii < 47 Or KeyAscii > 58 Then
            KeyAscii = 0
        End If
    End If
End Sub
