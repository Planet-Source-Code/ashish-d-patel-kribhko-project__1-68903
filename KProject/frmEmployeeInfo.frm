VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEmployeeInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employee Information"
   ClientHeight    =   7110
   ClientLeft      =   1575
   ClientTop       =   735
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   7995
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
      Height          =   405
      Left            =   4020
      TabIndex        =   74
      Top             =   6570
      Width           =   1005
   End
   Begin VB.CommandButton cmdUPdate 
      Caption         =   "&Update"
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
      Left            =   2970
      TabIndex        =   73
      Top             =   6570
      Width           =   1005
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6105
      Left            =   180
      TabIndex        =   0
      Top             =   270
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
      TabPicture(0)   =   "frmEmployeeInfo.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtCode"
      Tab(0).Control(1)=   "txtId"
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(4)=   "Label34"
      Tab(0).Control(5)=   "Label32"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Contact and Office Details"
      TabPicture(1)   =   "frmEmployeeInfo.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.TextBox txtCode 
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
         Left            =   -68910
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   660
         Width           =   1185
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
         TabIndex        =   69
         Top             =   570
         Width           =   1185
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
         TabIndex        =   58
         Top             =   3720
         Width           =   7275
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
            TabIndex        =   63
            Top             =   300
            Width           =   6075
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
            TabIndex        =   62
            Top             =   720
            Width           =   6075
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
            TabIndex        =   61
            Top             =   1140
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
            TabIndex        =   60
            Top             =   1140
            Width           =   1635
         End
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
            TabIndex        =   59
            Top             =   1560
            Width           =   1965
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
            TabIndex        =   68
            Top             =   360
            Width           =   915
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
            TabIndex        =   67
            Top             =   780
            Width           =   915
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
            TabIndex        =   66
            Top             =   1200
            Width           =   375
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
            TabIndex        =   65
            Top             =   1200
            Width           =   795
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
            TabIndex        =   64
            Top             =   1620
            Width           =   510
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
         TabIndex        =   38
         Top             =   1140
         Width           =   7275
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
            TabIndex        =   48
            Top             =   2160
            Width           =   1995
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
            TabIndex        =   47
            Top             =   1740
            Width           =   1995
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
            TabIndex        =   46
            Top             =   1740
            Width           =   1635
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
            TabIndex        =   45
            Top             =   1320
            Width           =   6075
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
            TabIndex        =   44
            Top             =   900
            Width           =   6075
         End
         Begin VB.TextBox txtLname 
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
            Left            =   5505
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   480
            Width           =   1635
         End
         Begin VB.TextBox txtMname 
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
            Left            =   3795
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   480
            Width           =   1635
         End
         Begin VB.TextBox txtFname 
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
            Left            =   2085
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   480
            Width           =   1635
         End
         Begin VB.ComboBox cmbSalutation 
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
            ItemData        =   "frmEmployeeInfo.frx":0038
            Left            =   1065
            List            =   "frmEmployeeInfo.frx":0048
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   480
            Width           =   975
         End
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
            TabIndex        =   39
            Top             =   2160
            Width           =   2865
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
            TabIndex        =   57
            Top             =   540
            Width           =   600
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
            TabIndex        =   56
            Top             =   210
            Width           =   1635
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
            TabIndex        =   55
            Top             =   210
            Width           =   1635
            WordWrap        =   -1  'True
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
            TabIndex        =   54
            Top             =   210
            Width           =   1605
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
            TabIndex        =   53
            Top             =   960
            Width           =   915
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
            TabIndex        =   52
            Top             =   1380
            Width           =   915
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
            TabIndex        =   51
            Top             =   1800
            Width           =   375
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
            TabIndex        =   50
            Top             =   2220
            Width           =   510
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
            TabIndex        =   49
            Top             =   1800
            Width           =   795
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
         TabIndex        =   24
         Top             =   810
         Width           =   7275
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
            TabIndex        =   30
            Top             =   420
            Width           =   705
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
            TabIndex        =   29
            Top             =   420
            Width           =   1605
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
            TabIndex        =   28
            Top             =   840
            Width           =   2385
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
            TabIndex        =   27
            Top             =   1260
            Width           =   2385
         End
         Begin VB.ComboBox cmbGender 
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
            ItemData        =   "frmEmployeeInfo.frx":0063
            Left            =   5700
            List            =   "frmEmployeeInfo.frx":006D
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   420
            Width           =   1035
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
            ItemData        =   "frmEmployeeInfo.frx":007F
            Left            =   5700
            List            =   "frmEmployeeInfo.frx":009B
            TabIndex        =   25
            Top             =   1290
            Width           =   885
         End
         Begin MSComCtl2.DTPicker dtpDOB 
            Height          =   375
            Left            =   5700
            TabIndex        =   31
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
            Format          =   19791873
            CurrentDate     =   38820
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
            TabIndex        =   37
            Top             =   450
            Width           =   1230
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
            TabIndex        =   36
            Top             =   870
            Width           =   1020
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
            TabIndex        =   35
            Top             =   1320
            Width           =   555
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
            TabIndex        =   34
            Top             =   480
            Width           =   720
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
            TabIndex        =   33
            Top             =   930
            Width           =   480
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
            TabIndex        =   32
            Top             =   1320
            Width           =   1185
         End
      End
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
         TabIndex        =   1
         Top             =   2580
         Width           =   7275
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
            TabIndex        =   13
            Top             =   1590
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
            TabIndex        =   12
            Top             =   1170
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
            TabIndex        =   11
            Top             =   750
            Width           =   1650
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
            TabIndex        =   10
            Top             =   330
            Width           =   1650
         End
         Begin VB.TextBox txtBasic 
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
            TabIndex        =   9
            Top             =   2400
            Width           =   1395
         End
         Begin VB.TextBox txtHRA 
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
            Left            =   5700
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   300
            Width           =   1455
         End
         Begin VB.TextBox txtOtherAllowance 
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
            Left            =   5700
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   720
            Width           =   1455
         End
         Begin VB.ComboBox cmbPF 
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
            ItemData        =   "frmEmployeeInfo.frx":00C1
            Left            =   5700
            List            =   "frmEmployeeInfo.frx":00CB
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   1140
            Width           =   1245
         End
         Begin VB.CommandButton cmdDepartmentList 
            Caption         =   "List..."
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
            Left            =   3300
            TabIndex        =   5
            Top             =   330
            Width           =   735
         End
         Begin VB.CommandButton cmdSubDepartmentList 
            Caption         =   "List..."
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
            Left            =   3300
            TabIndex        =   4
            Top             =   750
            Width           =   735
         End
         Begin VB.CommandButton cmdDesignationList 
            Caption         =   "List..."
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
            Left            =   3300
            TabIndex        =   3
            Top             =   1170
            Width           =   735
         End
         Begin VB.CommandButton cmdShiftList 
            Caption         =   "List..."
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
            Left            =   3300
            TabIndex        =   2
            Top             =   1590
            Width           =   735
         End
         Begin MSComCtl2.DTPicker dtpJoinDate 
            Height          =   345
            Left            =   1590
            TabIndex        =   14
            Top             =   2010
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   19791873
            CurrentDate     =   38820
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
            TabIndex        =   23
            Top             =   390
            Width           =   1095
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
            TabIndex        =   22
            Top             =   780
            Width           =   1425
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
            TabIndex        =   21
            Top             =   1200
            Width           =   1125
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
            TabIndex        =   20
            Top             =   1620
            Width           =   1020
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
            TabIndex        =   19
            Top             =   2040
            Width           =   900
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
            TabIndex        =   18
            Top             =   2430
            Width           =   555
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
            TabIndex        =   17
            Top             =   360
            Width           =   480
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
            TabIndex        =   16
            Top             =   750
            Width           =   1500
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
            TabIndex        =   15
            Top             =   1200
            Width           =   300
         End
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
         TabIndex        =   72
         Top             =   630
         Width           =   240
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
         TabIndex        =   71
         Top             =   720
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmEmployeeInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsInformation As New ADODB.Recordset
Dim iDeptId As Integer, iSubDeptId As Integer, iDesgId As Integer, iShiftID As Integer

Private Sub cmbBloodGroup_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
    Call Form_Load
End Sub

Private Sub cmdUPdate_Click()
On Error GoTo UpdateError
Dim strSaveSql As String, cnUpdate As New ADODB.Connection

With cnUpdate
    .ConnectionString = gStrConnectionString
    .Open
    .BeginTrans
    .Execute "Delete from m_employee where emp_id=" & txtId & " and emp_code='" & txtCode & "'"
     strSaveSql = "Insert into m_Employee values (" & txtId & ",'" & txtCode & _
        "','" & cmbSalutation.Text & "','" & txtFname & "','" & txtMname & "','" & _
        txtLname & "','" & txtAdd1 & "','" & txtAdd2 & "','" & cmbCity.Text & "','" & _
        cmbState.Text & "','" & txtPincode & "','" & txtPermAdd1 & "','" & txtPermAdd2 & _
        "','" & cmbPermCity.Text & "','" & cmbPermState.Text & "','" & txtPermPincode & _
        "','" & cmbGender.Text & "',#" & dtpDOB.Value & "#,'" & cmbBloodGroup.Text & _
        "','" & txtStdCode & "','" & txtHomePhone & "','" & txtMobile & "','" & txtEmail & _
        "'," & iDeptId & "," & iSubDeptId & "," & iDesgId & _
        "," & iShiftID & ",#" & dtpJoinDate.Value & "#,'" & Left(txtFname, 1) & _
        "'," & Val(txtBasic) & "," & Val(txtHRA) & "," & Val(txtOtherAllowance) & ",'" & cmbPF.Text & "',1)"
    .Execute strSaveSql
    .CommitTrans
End With

MsgBox "Information Updated", vbInformation, "Update"
Unload Me

Exit Sub
UpdateError:
    MsgBox "Error: " & Err.Number & vbCrLf & "Desc :" & Err.Description, vbCritical, "Update Error"
    Err.Clear
End Sub

Private Sub Form_Load()
On Error GoTo LoadError
    Call gFormCenter(Me)
    Call mLockControl(False)
    With rsInformation
            .Open "select * from view_employee where emp_code='" & gStrUser & "'", _
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
            iDeptId = .Fields("Dept_id")
            iDesgId = .Fields("desg_id")
            iSubDeptId = .Fields("sub_dept_id")
            iShiftID = .Fields("shf_id")
        End With
    rsInformation.Close
    Call mLockControl(True)
Exit Sub
LoadError:
    MsgBox "Error: " & Err.Number & vbCrLf & "Desc: " & Err.Description, vbCritical, "Load Error"
    Err.Clear
End Sub

Public Sub mLockControl(blLock As Boolean)
    txtId.Locked = blLock
    txtCode.Locked = blLock
    cmbSalutation.Locked = blLock
    txtFname.Locked = blLock
    txtMname.Locked = blLock
    txtLname.Locked = blLock
    cmbGender.Locked = blLock
    txtDepartment.Locked = blLock
    txtSubDepartment.Locked = blLock
    txtDesignation.Locked = blLock
    txtShift.Locked = blLock
    txtBasic.Locked = blLock
    txtHRA.Locked = blLock
    txtOtherAllowance.Locked = blLock
    cmbPF.Locked = blLock
End Sub
