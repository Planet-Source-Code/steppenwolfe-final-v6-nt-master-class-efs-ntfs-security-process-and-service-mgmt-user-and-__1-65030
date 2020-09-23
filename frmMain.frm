VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NMC V1.6 - Test Harness"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   10800
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   10800
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picItems 
      Height          =   5655
      Index           =   5
      Left            =   0
      ScaleHeight     =   5595
      ScaleWidth      =   10695
      TabIndex        =   11
      Top             =   0
      Width           =   10755
      Begin VB.TextBox txtEFS 
         Height          =   4335
         Left            =   5520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   68
         Tag             =   "NTRQ"
         Top             =   210
         Width           =   4995
      End
      Begin VB.CommandButton cmdEFSDemo 
         Caption         =   "Create a Test Account"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   3270
         TabIndex        =   67
         Tag             =   "NTRQ"
         Top             =   510
         Width           =   2115
      End
      Begin VB.CommandButton cmdEFSDemo 
         Caption         =   "Encrypt a File"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   3300
         TabIndex        =   66
         Tag             =   "NTRQ"
         Top             =   1650
         Width           =   2115
      End
      Begin VB.CommandButton cmdEFSDemo 
         Caption         =   "Impersonate and View"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   2
         Left            =   3300
         TabIndex        =   65
         Tag             =   "NTRQ"
         Top             =   2790
         Width           =   2115
      End
      Begin VB.CommandButton cmdEFSDemo 
         Caption         =   "Revert and Decrypt"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   3
         Left            =   3300
         TabIndex        =   64
         Tag             =   "NTRQ"
         Top             =   3960
         Width           =   2115
      End
      Begin VB.CommandButton cmdEFSControls 
         Caption         =   "Select a File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   210
         TabIndex        =   63
         Tag             =   "NTRQ"
         Top             =   5010
         Width           =   1695
      End
      Begin VB.CommandButton cmdEFSControls 
         Caption         =   "Select a Folder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   62
         Tag             =   "NTRQ"
         Top             =   5010
         Width           =   1695
      End
      Begin VB.CommandButton cmdEFSControls 
         Caption         =   "Encrypt "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3870
         TabIndex        =   61
         Tag             =   "NTRQ"
         Top             =   5010
         Width           =   1695
      End
      Begin VB.CommandButton cmdEFSControls 
         Caption         =   "Decrypt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   5700
         TabIndex        =   60
         Tag             =   "NTRQ"
         Top             =   5010
         Width           =   1695
      End
      Begin VB.TextBox txtEFSPath 
         Height          =   315
         Left            =   7590
         TabIndex        =   59
         Tag             =   "NTRQ"
         Top             =   5070
         Width           =   2775
      End
      Begin MSComDlg.CommonDialog cdFile 
         Left            =   9600
         Top             =   3840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblEFS 
         BackStyle       =   0  'Transparent
         Caption         =   "Create a new folder, add a test user account, and copy the test file to the folder."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   945
         Index           =   0
         Left            =   450
         TabIndex        =   74
         Top             =   510
         Width           =   2805
      End
      Begin VB.Label lblEFS 
         BackStyle       =   0  'Transparent
         Caption         =   "Encrypt the test file 'sample.txt'. You can still access it normally with your account."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   945
         Index           =   1
         Left            =   480
         TabIndex        =   73
         Top             =   1650
         Width           =   2805
      End
      Begin VB.Label lblEFS 
         BackStyle       =   0  'Transparent
         Caption         =   "Log on the new User account and attempt to view the sample file."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   945
         Index           =   2
         Left            =   480
         TabIndex        =   72
         Top             =   2850
         Width           =   2805
      End
      Begin VB.Label lblEFS 
         BackStyle       =   0  'Transparent
         Caption         =   "Revert security token back to your account. Decrypt and Delete the file."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   945
         Index           =   3
         Left            =   480
         TabIndex        =   71
         Top             =   3990
         Width           =   2805
      End
      Begin VB.Label lblEFSStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NTFS Drives:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   5550
         TabIndex        =   70
         Top             =   4590
         Width           =   1020
      End
      Begin VB.Label lblEFSStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Path:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   7620
         TabIndex        =   69
         Top             =   4860
         Width           =   405
      End
      Begin VB.Label lblItems 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Encrypted File System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   1875
      End
   End
   Begin VB.PictureBox picItems 
      Height          =   5625
      Index           =   4
      Left            =   0
      ScaleHeight     =   5565
      ScaleWidth      =   10695
      TabIndex        =   9
      Top             =   0
      Width           =   10755
      Begin VB.CommandButton cmdRegDemo 
         Caption         =   "Cleanup >>"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   1260
         TabIndex        =   102
         Top             =   3150
         Width           =   1965
      End
      Begin VB.CommandButton cmdRegDemo 
         Caption         =   "Test Security >>"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   1260
         TabIndex        =   101
         Top             =   2130
         Width           =   1965
      End
      Begin VB.CommandButton cmdRegDemo 
         Caption         =   "Start >>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   1230
         TabIndex        =   100
         Top             =   1110
         Width           =   1965
      End
      Begin VB.TextBox txtRegPath 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   270
         Locked          =   -1  'True
         TabIndex        =   99
         Text            =   "HKEY_CURRENT_USER\"
         Top             =   4530
         Width           =   4425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Test Key:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   270
         TabIndex        =   106
         Top             =   4320
         Width           =   765
      End
      Begin VB.Label lblRegDemo 
         Appearance      =   0  'Flat
         Caption         =   "Resets the key security. Deletes the key. Deletes the temporary User account."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   765
         Index           =   2
         Left            =   3540
         TabIndex        =   105
         Top             =   3150
         Width           =   4455
      End
      Begin VB.Label lblRegDemo 
         Appearance      =   0  'Flat
         Caption         =   "Sets new key with 'Read Only' rights for the new account. Launches Regedit so you can check the permissions change."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   765
         Index           =   1
         Left            =   3540
         TabIndex        =   104
         Top             =   2130
         Width           =   4455
      End
      Begin VB.Label lblRegDemo 
         Appearance      =   0  'Flat
         Caption         =   "Creates a User Account Reg_Demo, adds the key Software\NMC Reg Demo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   795
         Index           =   0
         Left            =   3540
         TabIndex        =   103
         Top             =   1110
         Width           =   4455
      End
      Begin VB.Label lblEFSStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NTFS Drives:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   720
         TabIndex        =   75
         Top             =   5370
         Width           =   1020
      End
      Begin VB.Label lblItems 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registry Security"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.PictureBox picItems 
      Height          =   5745
      Index           =   3
      Left            =   0
      ScaleHeight     =   5685
      ScaleWidth      =   10695
      TabIndex        =   7
      Top             =   0
      Width           =   10755
      Begin VB.Frame fmNTFS 
         Caption         =   "Permissions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1725
         Index           =   0
         Left            =   360
         TabIndex        =   93
         Top             =   360
         Width           =   2625
         Begin VB.CheckBox chkPermissions 
            Caption         =   "Read Only"
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   98
            Top             =   300
            Value           =   1  'Checked
            Width           =   1125
         End
         Begin VB.CheckBox chkPermissions 
            Caption         =   "List Directories"
            Height          =   225
            Index           =   1
            Left            =   150
            TabIndex        =   97
            Top             =   570
            Width           =   1455
         End
         Begin VB.CheckBox chkPermissions 
            Caption         =   "Execute"
            Height          =   225
            Index           =   2
            Left            =   150
            TabIndex        =   96
            Top             =   840
            Width           =   975
         End
         Begin VB.CheckBox chkPermissions 
            Caption         =   "Write"
            Height          =   225
            Index           =   3
            Left            =   150
            TabIndex        =   95
            Top             =   1110
            Width           =   795
         End
         Begin VB.CheckBox chkPermissions 
            Caption         =   "Change"
            Height          =   225
            Index           =   4
            Left            =   150
            TabIndex        =   94
            Top             =   1380
            Width           =   945
         End
      End
      Begin VB.Frame fmNTFS 
         Caption         =   "Inheritence"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Index           =   1
         Left            =   360
         TabIndex        =   89
         Top             =   2160
         Width           =   2625
         Begin VB.OptionButton optInherit 
            Caption         =   "Subfolders and Children"
            Height          =   225
            Index           =   2
            Left            =   120
            TabIndex        =   92
            Top             =   840
            Width           =   2145
         End
         Begin VB.OptionButton optInherit 
            Caption         =   "Subfolders"
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   91
            Top             =   570
            Width           =   1245
         End
         Begin VB.OptionButton optInherit 
            Caption         =   "No Inheritance"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   90
            Top             =   300
            Value           =   -1  'True
            Width           =   1485
         End
      End
      Begin VB.Frame fmNTFS 
         Caption         =   "Recursive"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Index           =   2
         Left            =   360
         TabIndex        =   86
         Top             =   3420
         Width           =   2625
         Begin VB.OptionButton optRecurse 
            Caption         =   "Sub Folders and Children"
            Height          =   225
            Index           =   2
            Left            =   150
            TabIndex        =   107
            Top             =   840
            Width           =   2175
         End
         Begin VB.OptionButton optRecurse 
            Caption         =   "Apply to Sub Folders"
            Height          =   225
            Index           =   1
            Left            =   150
            TabIndex        =   88
            Top             =   540
            Width           =   2175
         End
         Begin VB.OptionButton optRecurse 
            Caption         =   "Apply to Parent Only"
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   87
            Top             =   270
            Value           =   -1  'True
            Width           =   2175
         End
      End
      Begin VB.Frame fmNTFS 
         Caption         =   "File/Folder Demo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3075
         Index           =   3
         Left            =   3180
         TabIndex        =   79
         Top             =   360
         Width           =   4665
         Begin VB.TextBox txtNTFSPath 
            Height          =   315
            Left            =   150
            TabIndex        =   84
            Tag             =   "NTRQ"
            Top             =   2100
            Width           =   4305
         End
         Begin VB.CommandButton cmdNTFSControls 
            Caption         =   "Select a Folder"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1950
            TabIndex        =   83
            Tag             =   "NTRQ"
            Top             =   1440
            Width           =   1695
         End
         Begin VB.CommandButton cmdNTFSControls 
            Caption         =   "Select a File"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   180
            TabIndex        =   82
            Tag             =   "NTRQ"
            Top             =   1440
            Width           =   1695
         End
         Begin VB.CommandButton cmdNTFSControls 
            Caption         =   "Commit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   2790
            TabIndex        =   81
            Tag             =   "NTRQ"
            Top             =   2490
            Width           =   1695
         End
         Begin VB.DirListBox dirNTFS 
            Height          =   1050
            Left            =   180
            TabIndex        =   80
            Top             =   330
            Width           =   4275
         End
         Begin VB.Label lblEFSStatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Path:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   180
            TabIndex        =   85
            Top             =   1920
            Width           =   405
         End
      End
      Begin VB.Frame fmNTFS 
         Caption         =   "Access Flag"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Index           =   5
         Left            =   360
         TabIndex        =   76
         Top             =   4620
         Width           =   2625
         Begin VB.OptionButton optAccess 
            Caption         =   "Permit Access"
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   78
            Top             =   270
            Value           =   -1  'True
            Width           =   1485
         End
         Begin VB.OptionButton optAccess 
            Caption         =   "Deny Access"
            Height          =   225
            Index           =   1
            Left            =   150
            TabIndex        =   77
            Top             =   540
            Width           =   1425
         End
      End
      Begin VB.Label lblItems 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NTFS Security"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   1185
      End
   End
   Begin VB.PictureBox picItems 
      Height          =   5685
      Index           =   2
      Left            =   0
      ScaleHeight     =   5625
      ScaleWidth      =   10695
      TabIndex        =   5
      Top             =   0
      Width           =   10755
      Begin VB.CommandButton cmdPrcControls 
         Caption         =   "Kill Thread"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   9150
         TabIndex        =   113
         Top             =   5190
         Width           =   1425
      End
      Begin VB.CommandButton cmdPrcControls 
         Caption         =   "Set Priority"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   7650
         TabIndex        =   112
         Top             =   5190
         Width           =   1425
      End
      Begin VB.CommandButton cmdPrcControls 
         Caption         =   "Get Priority"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   6150
         TabIndex        =   111
         Top             =   5190
         Width           =   1425
      End
      Begin VB.CommandButton cmdPrcControls 
         Caption         =   "Set Class"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   6150
         TabIndex        =   110
         Top             =   4740
         Width           =   1425
      End
      Begin VB.CommandButton cmdPrcControls 
         Caption         =   "Get Class"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   4650
         TabIndex        =   109
         Top             =   4740
         Width           =   1425
      End
      Begin VB.CommandButton cmdVerify 
         Caption         =   "Open Taskmgr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   8730
         TabIndex        =   58
         Top             =   4650
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cmdPrcControls 
         Caption         =   "Process Enum"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   54
         Top             =   4740
         Width           =   1425
      End
      Begin VB.CommandButton cmdPrcControls 
         Caption         =   "Process Exists"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1620
         TabIndex        =   53
         Top             =   4740
         Width           =   1425
      End
      Begin VB.CommandButton cmdPrcControls 
         Caption         =   "Kill Process"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3150
         TabIndex        =   52
         Top             =   4740
         Width           =   1425
      End
      Begin VB.CommandButton cmdPrcControls 
         Caption         =   "Module Enum"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   51
         Top             =   5190
         Width           =   1425
      End
      Begin VB.CommandButton cmdPrcControls 
         Caption         =   "Thread Enum"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   1620
         TabIndex        =   50
         Top             =   5190
         Width           =   1425
      End
      Begin VB.CommandButton cmdPrcControls 
         Caption         =   "Thread Susp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   3150
         TabIndex        =   49
         Top             =   5190
         Width           =   1425
      End
      Begin VB.CommandButton cmdPrcControls 
         Caption         =   "Thread Resm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   4650
         TabIndex        =   48
         Top             =   5190
         Width           =   1425
      End
      Begin MSComctlLib.ListView lstProcess 
         Height          =   4095
         Left            =   120
         TabIndex        =   55
         Top             =   480
         Width           =   10425
         _ExtentX        =   18389
         _ExtentY        =   7223
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblItems 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Process Management"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1875
      End
   End
   Begin VB.PictureBox picItems 
      Height          =   5715
      Index           =   1
      Left            =   0
      ScaleHeight     =   5655
      ScaleWidth      =   10695
      TabIndex        =   3
      Top             =   30
      Width           =   10755
      Begin VB.CommandButton cmdVerify 
         Caption         =   "Open Services.msc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   57
         Top             =   5100
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cmdSvcControls 
         Caption         =   "Enumerate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   4530
         Width           =   1365
      End
      Begin VB.CommandButton cmdSvcControls 
         Caption         =   "Start Svc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   1650
         TabIndex        =   21
         Top             =   4530
         Width           =   1365
      End
      Begin VB.CommandButton cmdSvcControls 
         Caption         =   "Stop Svc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   2
         Left            =   3180
         TabIndex        =   20
         Top             =   4530
         Width           =   1365
      End
      Begin VB.CommandButton cmdSvcControls 
         Caption         =   "Pause Svc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   3
         Left            =   4710
         TabIndex        =   19
         Top             =   4530
         Width           =   1365
      End
      Begin VB.CommandButton cmdSvcControls 
         Caption         =   "Restart Svc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   4
         Left            =   6180
         TabIndex        =   18
         Top             =   4530
         Width           =   1365
      End
      Begin VB.CommandButton cmdSvcControls 
         Caption         =   "Type Change"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   5
         Left            =   7680
         TabIndex        =   17
         Top             =   4530
         Width           =   1365
      End
      Begin VB.OptionButton optStartup 
         Caption         =   "Auto"
         Height          =   195
         Index           =   0
         Left            =   9240
         TabIndex        =   16
         Top             =   4530
         Width           =   825
      End
      Begin VB.OptionButton optStartup 
         Caption         =   "Manual"
         Height          =   195
         Index           =   1
         Left            =   9240
         TabIndex        =   15
         Top             =   4770
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton optStartup 
         Caption         =   "Disable"
         Height          =   195
         Index           =   2
         Left            =   9240
         TabIndex        =   14
         Top             =   5010
         Width           =   855
      End
      Begin MSComctlLib.ListView lstServices 
         Height          =   4305
         Left            =   90
         TabIndex        =   23
         Top             =   120
         Width           =   10545
         _ExtentX        =   18600
         _ExtentY        =   7594
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblItems 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Service Management"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   1800
      End
   End
   Begin VB.PictureBox picItems 
      Height          =   5775
      Index           =   0
      Left            =   0
      ScaleHeight     =   5715
      ScaleWidth      =   10695
      TabIndex        =   1
      Top             =   0
      Width           =   10755
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   420
         TabIndex        =   115
         Top             =   4500
         Width           =   1575
      End
      Begin VB.CommandButton cmdExtended 
         Caption         =   "*** Extended Example ***"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2460
         TabIndex        =   114
         Top             =   3090
         Width           =   2145
      End
      Begin VB.CommandButton cmdUsrControls 
         Caption         =   "Change Passwrd"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   9
         Left            =   330
         TabIndex        =   108
         Top             =   3660
         Width           =   1695
      End
      Begin VB.CommandButton cmdVerify 
         Caption         =   "Open Lusrmgr.msc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   4710
         TabIndex        =   56
         Top             =   5340
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cmdUsrControls 
         Caption         =   "Create a User"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   330
         TabIndex        =   36
         Top             =   690
         Width           =   1695
      End
      Begin VB.CommandButton cmdUsrControls 
         Caption         =   "User Exists?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   330
         TabIndex        =   35
         Top             =   1110
         Width           =   1695
      End
      Begin VB.CommandButton cmdUsrControls 
         Caption         =   "Create a Group"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   330
         TabIndex        =   34
         Top             =   1530
         Width           =   1695
      End
      Begin VB.CommandButton cmdUsrControls 
         Caption         =   "Add User to Group"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   330
         TabIndex        =   33
         Top             =   1950
         Width           =   1695
      End
      Begin VB.CommandButton cmdUsrControls 
         Caption         =   "Get User Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   330
         TabIndex        =   32
         Top             =   2370
         Width           =   1695
      End
      Begin VB.CommandButton cmdUsrControls 
         Caption         =   "Delete User"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   330
         TabIndex        =   31
         Top             =   2790
         Width           =   1695
      End
      Begin VB.CommandButton cmdUsrControls 
         Caption         =   "Delete Group"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   6
         Left            =   330
         TabIndex        =   30
         Top             =   3210
         Width           =   1695
      End
      Begin VB.TextBox txtUser 
         Height          =   315
         Index           =   0
         Left            =   4740
         TabIndex        =   29
         Text            =   "Test_User"
         Top             =   3120
         Width           =   3765
      End
      Begin VB.TextBox txtUser 
         Height          =   315
         Index           =   1
         Left            =   4740
         TabIndex        =   28
         Text            =   "Test_Pass"
         Top             =   3720
         Width           =   3765
      End
      Begin VB.TextBox txtUser 
         Height          =   315
         Index           =   2
         Left            =   4740
         TabIndex        =   27
         Top             =   4320
         Width           =   3765
      End
      Begin VB.TextBox txtUser 
         Height          =   315
         Index           =   3
         Left            =   4740
         TabIndex        =   26
         Text            =   "Test_Group"
         Top             =   4920
         Width           =   3765
      End
      Begin VB.CommandButton cmdUsrControls 
         Caption         =   "List Groups"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   7
         Left            =   3330
         TabIndex        =   25
         Top             =   1170
         Width           =   1305
      End
      Begin VB.CommandButton cmdUsrControls 
         Caption         =   "List Users"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   8
         Left            =   3330
         TabIndex        =   24
         Top             =   1560
         Width           =   1305
      End
      Begin MSComctlLib.ListView lstUsers 
         Height          =   2535
         Left            =   4680
         TabIndex        =   37
         Top             =   270
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   4471
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   210
         Index           =   0
         Left            =   330
         TabIndex        =   47
         Top             =   4200
         Width           =   45
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   210
         Index           =   1
         Left            =   330
         TabIndex        =   46
         Top             =   4440
         Width           =   45
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   210
         Index           =   2
         Left            =   330
         TabIndex        =   45
         Top             =   4680
         Width           =   45
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   210
         Index           =   3
         Left            =   330
         TabIndex        =   44
         Top             =   4920
         Width           =   45
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   210
         Index           =   4
         Left            =   330
         TabIndex        =   43
         Top             =   5160
         Width           =   45
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   210
         Index           =   5
         Left            =   330
         TabIndex        =   42
         Top             =   5400
         Width           =   45
      End
      Begin VB.Label lblDisplay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         Height          =   210
         Index           =   0
         Left            =   4740
         TabIndex        =   41
         Top             =   2910
         Width           =   840
      End
      Begin VB.Label lblDisplay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   210
         Index           =   1
         Left            =   4740
         TabIndex        =   40
         Top             =   3510
         Width           =   795
      End
      Begin VB.Label lblDisplay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Computer Name"
         Height          =   210
         Index           =   2
         Left            =   4740
         TabIndex        =   39
         Top             =   4110
         Width           =   1140
      End
      Begin VB.Label lblDisplay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group Name:"
         Height          =   210
         Index           =   3
         Left            =   4740
         TabIndex        =   38
         Top             =   4710
         Width           =   945
      End
      Begin VB.Label lblItems 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User and Group Management"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   2475
      End
   End
   Begin MSComctlLib.TabStrip tbItems 
      Height          =   6165
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   10874
      MultiRow        =   -1  'True
      Placement       =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Users and Groups"
            Key             =   "ug"
            Object.ToolTipText     =   "User and Group Management"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Service Management"
            Key             =   "sv"
            Object.ToolTipText     =   "System Service Management"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Process Management"
            Key             =   "pc"
            Object.ToolTipText     =   "System Process Management"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "NTFS Security"
            Key             =   "nt"
            Object.ToolTipText     =   "NTFS File System Security"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Registry Security"
            Key             =   "rg"
            Object.ToolTipText     =   "NTFS Registry Security"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "EFS Encryption"
            Key             =   "ef"
            Object.ToolTipText     =   "Encrypted File System"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   13
      Top             =   6315
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   609
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ~*** NMC V1.6 Test Harness ***~

Private Declare Sub Sleep Lib "Kernel32.dll" (ByVal dwMilliseconds As Long)

Private WithEvents cNMC         As clsNMC
Attribute cNMC.VB_VarHelpID = -1
Private m_NTFSDrive             As String
Private m_sSubKey               As String
Private m_cItem                 As ListItem


Private Sub cmdExtended_Click()
'/* run through a complete user installation

Dim lServState      As Long
Dim sPath           As String
        
    With cNMC
        MsgBox "This example will walk you through the steps to creating" + vbNewLine + _
        "a new user, complete with profile path, user key, group assignment," + vbNewLine + _
        "and profile settings. It will then launch another application as" + vbNewLine + _
        "the new user, demonstrating launching a process with alternate credentials.", _
        vbInformation, "Extended Demo"
        '/* if user doesn't already exist, create it
        If Not .User_Exist(.Computer_Name, "User_Extended_Test") Then
            If .User_Create(.Computer_Name, "User_Extended_Test", "password", App.Path, "Test") Then
                MsgBox "The User: User_Extended_Test was created successfully.", _
                    vbInformation, "Extended Demo"
            Else
                MsgBox "The User could not be created. Demo will exit", vbExclamation, "Extended Demo"
                Exit Sub
            End If
        End If
        
        '/* apply policies to the new account
        If .User_SetData(.Computer_Name, "User_Extended_Test", User_Normal Or User_PassNotRequired) Then
            MsgBox "Success! The User does not require a password.", vbInformation, "Extended Demo"
        Else
            MsgBox "Could not change the account properties.", vbExclamation, "Extended Demo"
        End If
        
        '/* create a new group and add the user
        If .Group_Create(.Computer_Name, "Extended_Group", "Extended Test Group") Then
            MsgBox "Success! The Group: Extended_Group was created successfully.", vbInformation, "Extended Demo"
        Else
            MsgBox "Could not create the Group: Extended_Group. Does it already exist?", _
                vbExclamation, "Extended Demo"
        End If
        
        '/* add the user to the group
        If .Group_Add(.Computer_Name, "User_Extended_Test", "Extended_Group") Then
            MsgBox "Success! The User was added to the Extended_Group successfully.", _
                vbInformation, "Extended Demo"
        Else
            MsgBox "Could not add the user.", vbExclamation, "Extended Demo"
        End If
        
        '/* impersonate the new user
        If .User_Impersonate("User_Extended_Test", "password", .Computer_Name) Then
            MsgBox "Success! You are now Impersonating the new user.", vbInformation, "Extended Demo"
        Else
            MsgBox "Could not Impersonate the new user account, demo will exit.", _
                vbExclamation, "Extended Demo"
            Exit Sub
        End If
        
        '/* create a profile directory
        sPath = Get_Folder(CSIDL_PROFILE)
        sPath = Left$(sPath, InStrRev(sPath, Chr$(92)))
        If .User_LoadProfile("User_Extended_Test", sPath) Then
            MsgBox "Success! A new profile path has been created.", vbInformation, "Extended Demo"
        Else
            MsgBox "Could not create the new profile path.", vbExclamation, "Extended Demo"
        End If
        
        '/* revert the process token
        If .User_Revert Then
            MsgBox "Success! Process token has reverted back to the original owner.", _
                vbInformation, "Extended Demo"
        Else
            MsgBox "Could not revert process token.", vbExclamation, "Extended Demo"
        End If
        
        '/* test for secondary logon process
        lServState = .Service_Query("Seclogon")
        If Not lServState = 4 Then
            '/* start the service
            If .Service_Start("seclogon") Then
                MsgBox "Success! Secondary logon service has started.", vbInformation, "Extended Demo"
            Else
                MsgBox "Could not start secondary logon process, exiting demo.", vbExclamation, "Extended Demo"
                Exit Sub
            End If
        End If
        
        '/* load outlook express
        sPath = Get_Folder(CSIDL_PROGRAM_FILES)
        If File_Exists(sPath + "\Outlook Express\msimn.exe") Then
            If .User_RunAs("User_Extended_Test", "password", _
                ".", sPath + "\Outlook Express\msimn.exe") Then
                MsgBox "Success! Outlook has started with the alternate user account." + vbNewLine + _
                    "Open Task and look at the user name that corresponds with Outlook Express.", _
                    vbInformation, "Extended Demo"
            Else
                MsgBox "Could not start application with Secondary Logon, exiting demo.", _
                    vbExclamation, "Extended Demo"
                Exit Sub
            End If
        '/* test for internet explorer
        ElseIf File_Exists(sPath + "\Internet Explorer\iexplore.exe") Then
            If .User_RunAs("User_Extended_Test", "password", _
                .Computer_Name, sPath + "\Internet Explorer\iexplore.exe") Then
                MsgBox "Success! Internet Explorer has started with the alternate user account." + vbNewLine + _
                    "Open Task and look at the user name that corresponds with Internet Explorer.", _
                    vbInformation, "Extended Demo"
            Else
                MsgBox "Could not start application with Secondary Logon, exiting demo.", _
                    vbExclamation, "Extended Demo"
                Exit Sub
            End If
        '/* load explorer
        Else
            sPath = Get_Folder(CSIDL_WINDOWS) + "\explorer.exe"
            If .User_RunAs("User_Extended_Test", "password", _
                .Computer_Name, sPath) Then
                MsgBox "Success! Explorer has started with the alternate user account." + vbNewLine + _
                    "Open Task and look at the user name that corresponds with Explorer.", _
                    vbInformation, "Extended Demo"
            Else
                MsgBox "Could not start application with Secondary Logon, exiting demo.", _
                    vbExclamation, "Extended Demo"
                Exit Sub
            End If
        End If
    End With
    
End Sub

Private Sub Form_Load()

    Set cNMC = New clsNMC
    With lstServices
        .View = lvwReport
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .AllowColumnReorder = True
        .ColumnHeaders.Add 1, , "Display Name", .Width / 3
        .ColumnHeaders.Add 2, , "Service Name", .Width / 3
        .ColumnHeaders.Add 3, , "Service Status", (.Width / 3) - 100
    End With
    
    txtUser(2).Text = cNMC.Computer_Name
    Get_Groups
    tbItems_Click
    NTFSDrive_Test
    m_sSubKey = "Software\NMC Reg Demo"
    txtRegPath.Text = txtRegPath.Text + m_sSubKey
    If OS_Check Then
        '/* toolhelp (kernal32) - *** still works in xp, btw
        Get_Processes
    Else
        '/* psapi - nt and up
        Get_ProcessesG2
    End If
    cmdSvcControls_Click 0

End Sub


'> Global Events
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub cNMC_eNComplete(ByVal sTask As String)
'/* task completed
    stBar.SimpleText = sTask
End Sub

Private Sub cNMC_eNErrorCond(ByVal sRoutine As String, ByVal sError As String)
'/* error condition
    MsgBox "The Routine: " + sRoutine + " could not complete.. Error: " + sError
End Sub


'> Service Management Routines
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub cmdSvcControls_Click(Index As Integer)
'/* service controls

Dim lItem   As ListItem

On Error GoTo Handler

    With cNMC
        Select Case Index
        '/* enumerate
        Case 0
            Dim cTemp As Collection
            Dim aData() As String
            Dim vItem As Variant
            Set cTemp = New Collection
            Set cTemp = .Service_Enumerate
            lstServices.ListItems.Clear
            For Each vItem In cTemp
                aData = Split(CStr(vItem), Chr$(31))
                With lstServices
                    Set lItem = .ListItems.Add(Text:=aData(0))
                    lItem.SubItems(1) = aData(1)
                    lItem.SubItems(2) = aData(2)
                End With
            Next vItem

        '/* start
        Case 1
            If .Service_Query(lstServices.SelectedItem.SubItems(1)) = 1 Then
                .Service_Start (lstServices.SelectedItem.SubItems(1))
            Else
                stBar.SimpleText = "Service is not in the Stopped State.. Aborting"
            End If
        
        '/* stop
        Case 2
            If .Service_Query(lstServices.SelectedItem.SubItems(1)) = 4 Then
                .Service_Stop (lstServices.SelectedItem.SubItems(1))
            Else
                stBar.SimpleText = "Service is not in the Started.. Aborting"
            End If
            
        '/* pause
        Case 3
            If .Service_Query(lstServices.SelectedItem.SubItems(1)) = 4 Then
                .Service_Pause (lstServices.SelectedItem.SubItems(1))
            Else
                stBar.SimpleText = "Service is not in the Started State.. Aborting"
            End If
        
        '/* resume
        Case 4
            If .Service_Query(lstServices.SelectedItem.SubItems(1)) < 3 Then
                .Service_Resume (lstServices.SelectedItem.SubItems(1))
            Else
                stBar.SimpleText = "Service is not in the Paused State.. Aborting"
            End If
            
        '/* change
        Case 5
            Select Case True
            '/* auto
            Case optStartup(0).Value
                .Service_Change lstServices.SelectedItem.SubItems(1), START_AUTO
            '/* manual
            Case optStartup(1).Value
                .Service_Change lstServices.SelectedItem.SubItems(1), START_DEMAND
            '/* disable
            Case optStartup(2).Value
                .Service_Change lstServices.SelectedItem.SubItems(1), START_DISABLED
            End Select
        End Select
    End With
    
    '/* refresh
    If Not Index = 0 Then
        Sleep 100
        Dim sOldText As String
        sOldText = stBar.SimpleText
        cmdSvcControls_Click 0
        stBar.SimpleText = sOldText
    End If
    
Handler:
    On Error GoTo 0

End Sub


'> User and Group Routines
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub cmdUsrControls_Click(Index As Integer)
'/* user/group controls

    With cNMC
        Select Case Index
        '/* create user
        Case 0
            If .User_Create(txtUser(2).Text, txtUser(0).Text, _
                txtUser(1).Text, App.Path, "test", App.Path + "\Scripts") Then
            End If
            Get_Users True
            
        '/* user exists
        Case 1
            .User_Exist txtUser(2).Text, txtUser(0).Text
            
        '/* create a group
        Case 2
            .Group_Create txtUser(2).Text, txtUser(3).Text, "Test Comment"
        
        '/* add to group
        Case 3
            .Group_Add txtUser(2).Text, txtUser(0).Text, txtUser(3).Text
            
        '/* user data
        Case 4
On Error Resume Next
            Dim cTemp As Collection
            Set cTemp = New Collection
            Set cTemp = .User_Data(txtUser(2).Text, txtUser(0).Text)
            lblUser(0) = cTemp.Item(1)
            lblUser(1) = cTemp.Item(2)
            lblUser(2) = cTemp.Item(3)
            lblUser(3) = cTemp.Item(4)
            lblUser(4) = cTemp.Item(5)
            lblUser(5) = cTemp.Item(6)
On Error GoTo 0
            
        '/* delete user
        Case 5
            .User_Delete txtUser(2).Text, txtUser(0).Text
            Get_Users True
            
        '/* delete group
        Case 6
            .Group_Delete txtUser(2).Text, txtUser(3).Text
            
        '/* list groups
        Case 7
            Get_Groups
            
        '/* list users
        Case 8
            Get_Users True
        
        '/* change password
        Case 9
            .User_Password .Computer_Name, txtUser(0).Text, "", "New_Pass1"
        End Select
    End With
    
End Sub

Private Sub Get_Users(ByVal bLocal As Boolean)
'/* get users list

Dim vItem       As Variant
Dim sUData()    As String

On Error Resume Next

    With lstUsers
        .ListItems.Clear
        .View = lvwReport
        .AllowColumnReorder = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "User Name", .Width - 100
    End With
    
    For Each vItem In cNMC.Users_List(txtUser(2).Text, "", bLocal)
        If Len(vItem) = 0 Then GoTo skip
        Set m_cItem = lstUsers.ListItems.Add(Text:=vItem)
skip:
    Next vItem

On Error GoTo 0

End Sub

Private Sub Get_Groups()
'/* get user group list

Dim vItem       As Variant
Dim sUData()    As String

    With lstUsers
        .ListItems.Clear
        .View = lvwReport
        .AllowColumnReorder = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "Group Name", .Width / 2
        .ColumnHeaders.Add 2, , "SID", .Width / 4
        .ColumnHeaders.Add 3, , "Access", (.Width / 4) - 100
    End With
    
    For Each vItem In cNMC.Groups_List
        sUData = Split(CStr(vItem), Chr$(31))
        If Len(sUData(0)) = 0 Then GoTo skip
        Set m_cItem = lstUsers.ListItems.Add(Text:=sUData(0))
        m_cItem.SubItems(1) = sUData(1)
        m_cItem.SubItems(2) = sUData(2)
skip:
    Next vItem
    
End Sub


'> Process Management Routines
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub cmdPrcControls_Click(Index As Integer)
'/* process controls

    With cNMC
        Select Case Index
        '/* process list
        Case 0
            If OS_Check Then
                '/* toolhelp (kernal32) - ***still works in xp
                Get_Processes
            Else
                '/* psapi - nt and up
                Get_ProcessesG2
            End If
            
        '/* process exists
        Case 1
            .Process_Exists "alg.exe", True
            
        '/* kill process
        Case 2
            .Process_Terminate lstProcess.SelectedItem.SubItems(1)
            
        '/* get process class
        Case 3
            If OS_Check Then
                If Not .Process_Exists(lstProcess.SelectedItem.Text, True) Then
                    MsgBox "Please select a valid process before continuing!", vbExclamation, "No Process Name!"
                    Exit Sub
                End If
            Else
                If Not .Process_ExistsG2(lstProcess.SelectedItem.Text, True) Then
                    MsgBox "Please select a valid process before continuing!", vbExclamation, "No Process Name!"
                    Exit Sub
                End If
            End If
            .Process_GetClass lstProcess.SelectedItem.SubItems(1)
        
        '/* set process class
        Case 4
            If OS_Check Then
                If Not .Process_Exists(lstProcess.SelectedItem.Text, True) Then
                    MsgBox "Please select a valid process before continuing!", vbExclamation, "No Process Name!"
                    Exit Sub
                End If
            Else
                If Not .Process_ExistsG2(lstProcess.SelectedItem.Text, True) Then
                    MsgBox "Please select a valid process before continuing!", vbExclamation, "No Process Name!"
                    Exit Sub
                End If
            End If
            If .Process_SetClass(lstProcess.SelectedItem.SubItems(1), Process_High) Then
                MsgBox lstProcess.SelectedItem + " has been changed to Class: High Priority.", _
                    vbInformation, "Process Class Change"
            End If
            
        '/* module list
        Case 5
            If OS_Check Then
                If Not .Process_Exists(lstProcess.SelectedItem.Text, True) Then
                    MsgBox "Please select a valid process before continuing!", vbExclamation, "No Process Name!"
                    Exit Sub
                End If
            Else
                If Not .Process_ExistsG2(lstProcess.SelectedItem.Text, True) Then
                    MsgBox "Please select a valid process before continuing!", vbExclamation, "No Process Name!"
                    Exit Sub
                End If
            End If
            If OS_Check Then
                '/* toolhelp (kernal32) - ***still works in xp
                Get_Modules lstProcess.SelectedItem.SubItems(1)
            Else
                '/* psapi - nt and up
                Get_ModulesG2 lstProcess.SelectedItem.SubItems(1)
            End If
            
        '/* thread list
        Case 6
            If OS_Check Then
                If Not .Process_Exists(lstProcess.SelectedItem.Text, True) Then
                    MsgBox "Please select a valid process before continuing!", vbExclamation, "No Process Name!"
                    Exit Sub
                End If
            Else
                If Not .Process_ExistsG2(lstProcess.SelectedItem.Text, True) Then
                    MsgBox "Please select a valid process before continuing!", vbExclamation, "No Process Name!"
                    Exit Sub
                End If
            End If
            Get_Threads lstProcess.SelectedItem.SubItems(1)
        
        '/* suspend thread
        Case 7
            If Not IsNumeric(lstProcess.SelectedItem.Text) Then
                MsgBox "Please select a valid Thread ID before continuing!", vbExclamation, "Invalid Input"
                Exit Sub
            End If
            .Thread_Suspend CLng(lstProcess.SelectedItem.Text), ""
            
        '/* resume thread
        Case 8
            If Not IsNumeric(lstProcess.SelectedItem.Text) Then
                MsgBox "Please select a valid Thread ID before continuing!", vbExclamation, "Invalid Input"
                Exit Sub
            End If
            .Thread_Resume CLng(lstProcess.SelectedItem.Text), ""
        
        '/* get thread priority
        Case 9
            If Not IsNumeric(lstProcess.SelectedItem.Text) Then
                MsgBox "Please select a valid Thread ID before continuing!", vbExclamation, "Invalid Input"
                Exit Sub
            End If
            .Thread_GetPriority lstProcess.SelectedItem.Text
            
        '/* set thread priority
        Case 10
            If Not IsNumeric(lstProcess.SelectedItem.Text) Then
                MsgBox "Please select a valid Thread ID before continuing!", vbExclamation, "Invalid Input"
                Exit Sub
            End If
            If .Thread_SetPriority(lstProcess.SelectedItem.Text, Thread_Maximum) Then
                MsgBox lstProcess.SelectedItem + " thread, has been changed to Maximum Priority.", _
                    vbInformation, "Thread Priority Change"
            End If
            
        '/* kill thread
        Case 11
            If Not IsNumeric(lstProcess.SelectedItem.Text) Then
                MsgBox "Please select a valid Thread ID before continuing!", vbExclamation, "Invalid Input"
                Exit Sub
            End If
            If .Thread_Terminate(lstProcess.SelectedItem.Text) Then
                MsgBox lstProcess.SelectedItem + " thread has been Terminated.", _
                    vbInformation, "Thread Terminated Successfully"
            End If
        
        End Select
    End With

End Sub

Private Sub Get_Processes()
'/* get process list

Dim vItem       As Variant
Dim cTemp       As Collection
Dim sUData()    As String

On Error Resume Next

    With lstProcess
        .View = lvwReport
        .LabelEdit = lvwManual
        .ListItems.Clear
        .ColumnHeaders.Clear
        .FullRowSelect = True
        .AllowColumnReorder = True
        .ColumnHeaders.Add 1, , "Name", (.Width / 7) * 3
        .ColumnHeaders.Add 2, , "ID", .Width / 7
        .ColumnHeaders.Add 3, , "Threads", .Width / 7
        .ColumnHeaders.Add 4, , "Priority", .Width / 7
        .ColumnHeaders.Add 5, , "Parent", .Width / 7
    End With
    
    Set cTemp = cNMC.Process_Enumerate
    For Each vItem In cTemp
        sUData = Split(CStr(vItem), Chr$(31))
        Set m_cItem = lstProcess.ListItems.Add(Text:=sUData(0))
        With m_cItem
            .SubItems(1) = sUData(1)
            .SubItems(2) = sUData(2)
            .SubItems(3) = sUData(3)
            .SubItems(4) = sUData(4)
        End With
    Next vItem

On Error GoTo 0

End Sub

Private Sub Get_ProcessesG2()
'/* get process list

Dim vItem       As Variant
Dim cTemp       As Collection
Dim sUData()    As String

On Error Resume Next

    With lstProcess
        .View = lvwReport
        .LabelEdit = lvwManual
        .ListItems.Clear
        .ColumnHeaders.Clear
        .FullRowSelect = True
        .AllowColumnReorder = True
        .ColumnHeaders.Add 1, , "Name", (.Width / 7) * 2
        .ColumnHeaders.Add 2, , "ID", .Width / 7
        .ColumnHeaders.Add 3, , "Path", (.Width / 7) * 4
    End With
    
    Set cTemp = cNMC.Process_EnumG2
    For Each vItem In cTemp
        sUData = Split(CStr(vItem), Chr$(31))
        Set m_cItem = lstProcess.ListItems.Add(Text:=sUData(0))
        With m_cItem
            .SubItems(1) = sUData(1)
            .SubItems(2) = sUData(2)
        End With
    Next vItem

On Error GoTo 0

End Sub

Private Sub Get_Modules(ByVal sProcess As String)
'/* get module list

Dim vItem       As Variant
Dim cTemp       As Collection
Dim sUData()    As String

On Error Resume Next

    '/* check for mods first
    Set cTemp = cNMC.Module_Enumerate(sProcess)
    If cTemp.Count = 0 Then Exit Sub
    
    With lstProcess
        .View = lvwReport
        .LabelEdit = lvwManual
        .ListItems.Clear
        .ColumnHeaders.Clear
        .FullRowSelect = True
        .AllowColumnReorder = True
        .ColumnHeaders.Add 1, , "Name", (.Width / 8) * 2
        .ColumnHeaders.Add 2, , "Path", .Width / 8
        .ColumnHeaders.Add 3, , "Usage", .Width / 8
        .ColumnHeaders.Add 4, , "Handle", .Width / 8
        .ColumnHeaders.Add 5, , "Address", .Width / 8
        .ColumnHeaders.Add 6, , "Proc", .Width / 8
        .ColumnHeaders.Add 7, , "Parent", .Width / 8
    End With
    
    For Each vItem In cTemp
        sUData = Split(CStr(vItem), Chr$(31))
        Set m_cItem = lstProcess.ListItems.Add(Text:=sUData(0))
        With m_cItem
            .SubItems(1) = sUData(1)
            .SubItems(2) = sUData(2)
            .SubItems(3) = sUData(3)
            .SubItems(4) = sUData(4)
            .SubItems(5) = sUData(5)
            .SubItems(6) = sUData(6)
        End With
    Next vItem

On Error GoTo 0

End Sub

Private Sub Get_ModulesG2(ByVal sProcess As String)
'/* get module list nt

Dim vItem       As Variant
Dim cTemp       As Collection
Dim sUData()    As String

On Error Resume Next

    '/* check for mods first
    Set cTemp = cNMC.Module_EnumG2(sProcess)
    If cTemp.Count = 0 Then Exit Sub
    
    With lstProcess
        .View = lvwReport
        .LabelEdit = lvwManual
        .ListItems.Clear
        .ColumnHeaders.Clear
        .FullRowSelect = True
        .AllowColumnReorder = True
        .ColumnHeaders.Add 1, , "Name", (.Width / 8) * 2
        .ColumnHeaders.Add 2, , "Path", (.Width / 8) * 4
        .ColumnHeaders.Add 3, , "Handle", .Width / 8
        .ColumnHeaders.Add 4, , "Parent ID", .Width / 8
    End With
    
    For Each vItem In cTemp
        sUData = Split(CStr(vItem), Chr$(31))
        Set m_cItem = lstProcess.ListItems.Add(Text:=sUData(0))
        With m_cItem
            .SubItems(1) = sUData(1)
            .SubItems(2) = sUData(2)
            .SubItems(3) = sUData(3)
            .SubItems(4) = sUData(4)
        End With
    Next vItem

On Error GoTo 0

End Sub

Private Sub Get_Threads(ByVal sProcess As String)
'/* get thread list

Dim vItem       As Variant
Dim cTemp       As Collection
Dim sUData()    As String

On Error Resume Next

    '/* check for mods first
    Set cTemp = cNMC.Thread_Enumerate(sProcess)
    If cTemp.Count = 0 Then Exit Sub
    
    With lstProcess
        .View = lvwReport
        .LabelEdit = lvwManual
        .ListItems.Clear
        .ColumnHeaders.Clear
        .FullRowSelect = True
        .AllowColumnReorder = True
        .ColumnHeaders.Add 1, , "ID", .Width / 7
        .ColumnHeaders.Add 2, , "Usage", .Width / 7
        .ColumnHeaders.Add 3, , "Flags", .Width / 7
        .ColumnHeaders.Add 4, , "Parent", .Width / 7
        .ColumnHeaders.Add 5, , "Base", .Width / 7
        .ColumnHeaders.Add 6, , "Delta", .Width / 7
    End With
    
    '/* put data to list
    For Each vItem In cTemp
        sUData = Split(CStr(vItem), Chr$(31))
        Set m_cItem = lstProcess.ListItems.Add(Text:=sUData(0))
        With m_cItem
            .SubItems(1) = sUData(1)
            .SubItems(2) = sUData(2)
            .SubItems(3) = sUData(3)
            .SubItems(4) = sUData(4)
            .SubItems(5) = sUData(5)
            .SubItems(6) = sUData(6)
        End With
    Next vItem

On Error GoTo 0

End Sub


'> EFS Routines
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub cmdEFSDemo_Click(Index As Integer)
'/* demo

Dim sPath       As String

    sPath = m_NTFSDrive + "NMC\sample.txt"
    
    With cNMC
        Select Case Index
        '/* setup
        Case 0
            '/* copy file to new directory
            CopyFile App.Path + "\sample.txt", sPath, 0
            '/* get text
            txtEFS.Text = File_Data(sPath)
            cmdEFSDemo(1).Enabled = True
            lblEFS(0).ForeColor = &H404040
            lblEFS(1).ForeColor = &HC00000
        
        '/* encrypt
        Case 1
            '/* clear textbox
            txtEFS.Text = ""
            '/* encrypt file
            If .EFS_Encrypt(sPath) Then
                Debug.Print "encryption success!"
            End If
            '/* reset text
            txtEFS.Text = File_Data(sPath)
            cmdEFSDemo(2).Enabled = True
            lblEFS(1).ForeColor = &H404040
            lblEFS(2).ForeColor = &HC00000
            
        '/* impersonate and view
        Case 2
            '/* test and create user account
            If Not .User_Exist(.Computer_Name, "EFS_Test") Then
                .User_Create .Computer_Name, "EFS_Test", "password", App.Path, "test", App.Path
            End If
            '/* impersonate new account
            If .User_Impersonate("EFS_Test", "password", .Computer_Name) Then
                Debug.Print "impersonate success!"
                txtEFS.Text = ""
            Else
                MsgBox "User: EFS_Test Password: password does not exist!" & vbNewLine & _
                "Please create the account manually.", _
                vbInformation, "Account does not exist!"
                Exit Sub
            End If
            '/* try to access file
            '/* should get a file access error
            txtEFS.Text = File_Data(sPath)
            cmdEFSDemo(3).Enabled = True
            lblEFS(2).ForeColor = &H404040
            lblEFS(3).ForeColor = &HC00000
        
        '/* revert token and reset
        Case 3
            '/* restore credentials
            If .User_Revert Then
                Debug.Print "account reverted!"
            End If
            '/* load text
            .EFS_Decrypt sPath
            txtEFS.Text = File_Data(sPath)
            '/* delete user account
            .User_Delete .Computer_Name, "EFS_Test"
        End Select
    End With
    
End Sub

Private Function File_Data(ByVal sPath As String) As String
'/* extract text from file

Dim lLen        As Long
Dim sText       As String

On Error GoTo Handler

    lLen = FileLen(sPath)
    sText = Space$(lLen)
    Open sPath For Binary As #1
    Get #1, , sText
    Close #1
    File_Data = sText
    
Exit Function

Handler:
    '/* get the error cond
    Debug.Print Err.Description + " Err# " + CStr(Err.Number)

End Function

Private Sub cmdEFSControls_Click(Index As Integer)
'/* EFS controls

    Select Case Index
    '/* select file
    Case 0
        Select_File
        
    '/* select folder
    Case 1
        txtEFSPath.Text = Get_Path
        
    '/* encrypt
    Case 2
        If Len(txtEFSPath.Text) = 0 Then
            MsgBox "Please select a File or Directory before proceeding!", vbExclamation, "Invalid Input"
            Exit Sub
        End If
        With cNMC
            If .NTFS_Drive(txtEFSPath.Text) Then
                .EFS_Encrypt txtEFSPath.Text
            Else
                MsgBox "This drive is not formatted with NTFS!", vbExclamation, "Not NTFS!"
            End If
        End With
        
    '/* decrypt
    Case 3
        If Len(txtEFSPath.Text) = 0 Then
            MsgBox "Please select a File or Directory before proceeding!", vbExclamation, "Invalid Input"
            Exit Sub
        End If
        With cNMC
            If .NTFS_Drive(txtEFSPath.Text) Then
                .EFS_Decrypt txtEFSPath.Text
            Else
                MsgBox "This drive is not formatted with NTFS!", vbExclamation, "Not NTFS!"
            End If
        End With
    End Select
    
End Sub

Private Sub Select_File()
'/* get the file

On Error GoTo Handler

    With cdFile
        .DialogTitle = "Select a File"
        .CancelError = True
        .DefaultExt = ".txt"
        .InitDir = Left$(App.Path, 3)
        .ShowOpen
        txtEFSPath.Text = .FileName
    End With

    If Len(txtEFSPath.Text) > 0 Then
        stBar.SimpleText = "File: " + txtEFSPath.Text + " selected.."
    End If
    
Handler:

End Sub


'> NTFS Routines
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub NTFSDrive_Test()
'/* test ntfs availability
'/* create test directory

    If cNMC.NTFS_Check Is Nothing Then
        MsgBox "None of your drives are formatted with the NTFS File System." & vbNewLine & _
        "NTFS related options have been Disabled.", vbExclamation, "No NTFS Drives Detected!"
        Disable_NTFS
        lblEFSStatus(0).Caption = "There are no NTFS Formatted Drives.."
        Exit Sub
    Else
        Dim vItem As Variant
        Dim sDrive As String
        For Each vItem In cNMC.NTFS_Check
            sDrive = sDrive + CStr(vItem)
            m_NTFSDrive = vItem
        Next vItem
        If m_NTFSDrive = "" Then
            Disable_NTFS
            m_NTFSDrive = "c:\"
        End If
        lblEFSStatus(0).Caption = "NTFS Drives: " + sDrive
        cNMC.Create_Directory m_NTFSDrive + "NMC\"
        txtEFSPath.Text = m_NTFSDrive + "NMC\"
        dirNTFS.Path = m_NTFSDrive + "NMC\"
        txtNTFSPath.Text = m_NTFSDrive + "NMC\"
    End If

End Sub

Private Sub Disable_NTFS()
'/* disable controls if NTFS not found

Dim oCtrl   As Control

On Error Resume Next

    For Each oCtrl In Controls
        If oCtrl.Tag = "NTRQ" Then
            oCtrl.Enabled = False
        End If
    Next oCtrl
    
On Error GoTo 0

End Sub

Private Sub cmdRegDemo_Click(Index As Integer)

    With cNMC
        Select Case Index
        '/* modify key permissions
        Case 0
            With New clsLightning
                .Create_Key HKEY_CURRENT_USER, m_sSubKey
                .Write_String HKEY_CURRENT_USER, m_sSubKey, "test", "NMC Test Value"
            End With
            If .User_Create(.Computer_Name, "Reg_Demo", "password", App.Path, "test", App.Path) Then
                MsgBox "The account Reg_Demo has been created!", vbInformation, "Success!"
                cmdRegDemo(1).Enabled = True
                lblRegDemo(0).ForeColor = &H404040
                lblRegDemo(1).ForeColor = &HC00000
            End If
            
        '/* modify security
        Case 1
            If .NTFS_Key(HKEY_CURRENT_USER, m_sSubKey, _
                "Reg_Demo", Registry_Read, _
                Access_Allowed, Non_Propogate) Then
                MsgBox "Read-Write permit access to the key: " + m_RegPath + vbNewLine + _
                " has been granted to the Reg_Demo account.", vbInformation, "Success!"
            End If
            If .Group_Add(.Computer_Name, "Reg_Demo", "Administrators") Then
                MsgBox "User Reg_Demo added to the Administrators group.", vbInformation, "Success!"
                MsgBox "Now check the key permissions for the " + vbNewLine + _
                "HKEY_CURRENT_USER\Software\NMC Reg Demo - key." + vbNewLine + _
                "Right click - permissions - security tab - Reg_Demo user.", _
                vbInformation, "Success!"
                Open_File "regedit.exe"
                cmdRegDemo(2).Enabled = True
                lblRegDemo(1).ForeColor = &H404040
                lblRegDemo(2).ForeColor = &HC00000
            End If
                
        '/* reset and delete
        Case 2
            If .NTFS_Key(HKEY_CURRENT_USER, m_RegPath, _
                "Administrator", Registry_Full_Control, _
                Access_Allowed, Non_Propogate) Then
                MsgBox "Full control has been restored to key: " + m_RegPath + vbNewLine + _
                " to the Reg_Demo account.", vbInformation, "Success!"
            End If
            With New clsLightning
                .Delete_Key HKEY_CURRENT_USER, m_RegPath
            End With
            If .User_Delete(.Computer_Name, "Reg_Demo") Then
                MsgBox "The RegDemo account has been deleted!", vbInformation, "Success!"
            End If
        End Select
    End With
    
End Sub

Private Sub cmdNTFSControls_Click(Index As Integer)

Dim lResult     As Long

    With cNMC
        Select Case Index
        '/* select file
        Case 0
            Select_File
        
        '/* select folder
        Case 1
            txtNTFSPath.Text = Get_Path
        
        '/* folder permissions
        Case 2
            If .NTFS_Drive(txtNTFSPath.Text) Then
                lResult = Folder_Access
                If lResult = -1 Then
                    MsgBox "This is not a valid Folder permissions set, Please try again", _
                    vbInformation, "Invalid Input!"
                    Exit Sub
                End If
                '/* sub folders
                If optRecurse(1).Value Then
                    .NTFS_Recursive txtNTFSPath.Text, "Administrator", _
                        Folder_Access, Inheritence_Flags, Access_Type
                '/* sub folders and children
                ElseIf optRecurse(2).Value Then
                    .NTFS_Recursive txtNTFSPath.Text, "Administrator", _
                        Folder_Access, Inheritence_Flags, Access_Type, True
                '/* parent folder
                Else
                    .NTFS_Folder txtNTFSPath.Text, "Administrator", _
                        Folder_Access, Inheritence_Flags, Access_Type
                End If
            End If
        End Select
    End With
    
End Sub

Private Function Folder_Access() As Long
'/* translate checkbox state to mask

    With chkPermissions
        Select Case 1
        Case .Item(4) And .Item(3) And .Item(2) And .Item(1) And .Item(0)
            Folder_Access = cNMC.Return_Folder(Folder_Full_Control)
        Case .Item(3) And .Item(2) And .Item(1) And .Item(0)
            Folder_Access = cNMC.Return_Folder(Folder_Read_Write_Execute_List)
        Case .Item(3) And .Item(2) And .Item(0)
            Folder_Access = cNMC.Return_Folder(Folder_Read_Write_Execute)
        Case .Item(3) And .Item(1) And .Item(0)
            Folder_Access = cNMC.Return_Folder(Folder_Read_Write_List)
        Case .Item(2) And .Item(1) And .Item(0)
            Folder_Access = cNMC.Return_Folder(Folder_Read_Execute_List)
        Case .Item(3) And .Item(0)
            Folder_Access = cNMC.Return_Folder(Folder_Read_Write)
        Case .Item(2) And .Item(0)
            Folder_Access = cNMC.Return_Folder(Folder_Read_Execute)
        Case .Item(1) And .Item(0)
            Folder_Access = cNMC.Return_Folder(Folder_Read_List)
        Case .Item(0)
            Folder_Access = cNMC.Return_Folder(Folder_Read)
        Case Else
            Folder_Access = -1
        End Select
    End With

End Function

Private Function Inheritence_Flags() As Long
'/* translate optionbox state to mask
    
    With cNMC
        Select Case True
        Case optInherit(0).Value:       Inheritence_Flags = .Return_Inherit(Non_Propogate)
        Case optInherit(1).Value:       Inheritence_Flags = .Return_Inherit(Container_Inherit)
        Case optInherit(2).Value:       Inheritence_Flags = .Return_Inherit(Child_Container_Inherit)
        End Select
    End With
    
End Function

Private Function Access_Type() As Long
'/* translate optionbox state to mask

    With cNMC
        Select Case True
        Case optAccess(0).Value:        Access_Type = .Return_Type(Access_Allowed)
        Case optAccess(1).Value:        Access_Type = .Return_Type(Access_Denied)
        End Select
    End With

End Function


Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub


'> Support Routines
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub tbItems_Click()
'/* tab controls

Dim oPic As PictureBox
    
    For Each oPic In picItems
        With oPic
            .Visible = False
            .BorderStyle = 0
            .Left = 0
            .Top = 0
            .Width = tbItems.Width
        End With
    Next oPic
    
    Select Case tbItems.SelectedItem.Index
    '/* users
    Case 1
        picItems(0).Visible = True
    '/* service
    Case 2
        picItems(1).Visible = True
    '/* process
    Case 3
        picItems(2).Visible = True
    '/* ntfs
    Case 4
        picItems(3).Visible = True
    '/* registry
    Case 5
        picItems(4).Visible = True
    '/* efs
    Case 6
        picItems(5).Visible = True
    End Select

End Sub

Public Sub Disable_Items()
'/* disable NT controls if 98/ME

Dim oCtrl As Control

On Error Resume Next

    For Each oCtrl In Controls
        If oCtrl.Tag = "UV" Then
            oCtrl.Enabled = False
        ElseIf oCtrl.Tag = "NTRQ" Then
            oCtrl.Enabled = False
        End If
    Next oCtrl

On Error GoTo 0

End Sub

Private Sub cmdVerify_Click(Index As Integer)

    Select Case Index
    '/* user manager
    Case 0
        Open_File m_SysPath + "lusrmgr.msc"
        
    '/* services
    Case 1
        Open_File m_SysPath + "services.msc"
    
    '/* task manager
    Case 2
        Open_File m_SysPath + "taskmgr.exe"
    End Select
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cNMC = Nothing
End Sub
