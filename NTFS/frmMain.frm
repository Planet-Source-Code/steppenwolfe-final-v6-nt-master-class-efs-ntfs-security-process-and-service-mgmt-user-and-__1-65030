VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NMC - NTFS Security"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10440
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
   ScaleHeight     =   5895
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
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
      Left            =   120
      TabIndex        =   26
      Top             =   3060
      Width           =   2625
      Begin VB.OptionButton optRecurse 
         Caption         =   "Apply to Parent Only"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   29
         Top             =   270
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton optRecurse 
         Caption         =   "Apply to Sub Folders"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   28
         Top             =   540
         Width           =   2175
      End
      Begin VB.OptionButton optRecurse 
         Caption         =   "Sub Folders and Children"
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   27
         Top             =   840
         Width           =   2175
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
      Left            =   120
      TabIndex        =   23
      Top             =   4320
      Width           =   2625
      Begin VB.OptionButton optAccess 
         Caption         =   "Permit Access"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   25
         Top             =   270
         Value           =   -1  'True
         Width           =   1485
      End
      Begin VB.OptionButton optAccess 
         Caption         =   "Deny Access"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   24
         Top             =   540
         Width           =   1425
      End
   End
   Begin VB.Frame fmNTFS 
      Caption         =   "Registry Demo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Index           =   4
      Left            =   2940
      TabIndex        =   13
      Top             =   3210
      Width           =   4695
      Begin VB.CommandButton cmdRegDemo 
         Caption         =   "Remove Key"
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
         Left            =   2280
         TabIndex        =   22
         Top             =   930
         Width           =   1965
      End
      Begin VB.CommandButton cmdRegDemo 
         Caption         =   "Apply to Key"
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
         Left            =   120
         TabIndex        =   21
         Top             =   930
         Width           =   1965
      End
      Begin VB.TextBox txtRegPath 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "HKEY_CURRENT_USER\"
         Top             =   510
         Width           =   4425
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
      Left            =   2940
      TabIndex        =   12
      Top             =   60
      Width           =   4665
      Begin VB.TextBox txtNTFSPath 
         Height          =   315
         Left            =   150
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
         Tag             =   "NTRQ"
         Top             =   2490
         Width           =   1695
      End
      Begin VB.DirListBox dirNTFS 
         Height          =   1050
         Left            =   180
         TabIndex        =   14
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
         Index           =   1
         Left            =   180
         TabIndex        =   19
         Top             =   1920
         Width           =   405
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
      Height          =   1155
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1860
      Width           =   2625
      Begin VB.OptionButton optInherit 
         Caption         =   "Subfolders and Children"
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   2145
      End
      Begin VB.OptionButton optInherit 
         Caption         =   "Subfolders"
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   570
         Width           =   1245
      End
      Begin VB.OptionButton optInherit 
         Caption         =   "No Inheritance"
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Value           =   -1  'True
         Width           =   1485
      End
   End
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
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   2625
      Begin VB.CheckBox chkPermissions 
         Caption         =   "Read Only"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   8
         Top             =   300
         Value           =   1  'Checked
         Width           =   1125
      End
      Begin VB.CheckBox chkPermissions 
         Caption         =   "List Directories"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   7
         Top             =   570
         Width           =   1455
      End
      Begin VB.CheckBox chkPermissions 
         Caption         =   "Execute"
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   6
         Top             =   840
         Width           =   975
      End
      Begin VB.CheckBox chkPermissions 
         Caption         =   "Write"
         Height          =   225
         Index           =   3
         Left            =   150
         TabIndex        =   5
         Top             =   1110
         Width           =   795
      End
      Begin VB.CheckBox chkPermissions 
         Caption         =   "Change"
         Height          =   225
         Index           =   4
         Left            =   150
         TabIndex        =   4
         Top             =   1380
         Width           =   945
      End
   End
   Begin MSComDlg.CommonDialog cdFile 
      Left            =   180
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   5550
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   609
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
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
      Left            =   150
      TabIndex        =   1
      Top             =   5280
      Width           =   1020
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cNTFS    As clsNTFS
Attribute cNTFS.VB_VarHelpID = -1
Private m_NTFSDrive         As String
Private m_sSubKey           As String

Private Sub Form_Load()
    
    Set cNTFS = New clsNTFS
    NTFSDrive_Test
    m_sSubKey = "Software\" + App.ProductName
    txtRegPath.Text = txtRegPath.Text + m_sSubKey

End Sub

Private Sub cNTFS_eNComplete(ByVal sTask As String)
'/* task completed
    stBar.SimpleText = sTask
End Sub

Private Sub cNTFS_eNErrorCond(ByVal sRoutine As String, ByVal sError As String)
'/* error condition
    MsgBox "The Routine: " + sRoutine + " could not complete.. Error: " + sError
End Sub

Private Sub cmdRegDemo_Click(Index As Integer)

    With cNTFS
        Select Case Index
        '/* modify key permissions
        Case 0
            With New clsLightning
                .Create_Key HKEY_CURRENT_USER, m_sSubKey
                .Write_String HKEY_CURRENT_USER, m_sSubKey, "test", "NMC Test Value"
            End With
            With cNTFS
                If .NTFS_Key(HKEY_CURRENT_USER, m_sSubKey, _
                    "Administrator", Registry_Read_Write, _
                    Access_Allowed, Non_Propogate) Then
                    MsgBox "Read-Write permit access to the key: " + m_sSubKey + vbNewLine + _
                    " has been granted to the Administrator account.", vbInformation, "Success!"
                End If
            End With
            
        '/* reset and delete
        Case 1
            With cNTFS
                If .NTFS_Key(HKEY_CURRENT_USER, m_sSubKey, _
                    "Administrator", Registry_Full_Control, _
                    Access_Allowed, Non_Propogate) Then
                    MsgBox "Full control has been restored to key: " + m_sSubKey + vbNewLine + _
                    " to the Administrator account.", vbInformation, "Success!"
                End If
            End With
            With New clsLightning
                .Delete_Key HKEY_CURRENT_USER, m_sSubKey
            End With
        End Select
    End With
    
End Sub

Private Sub cmdNTFSControls_Click(Index As Integer)

Dim lResult     As Long

    With cNTFS
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
            Folder_Access = cNTFS.Return_Folder(Folder_Full_Control)
        Case .Item(3) And .Item(2) And .Item(1) And .Item(0)
            Folder_Access = cNTFS.Return_Folder(Folder_Read_Write_Execute_List)
        Case .Item(3) And .Item(2) And .Item(0)
            Folder_Access = cNTFS.Return_Folder(Folder_Read_Write_Execute)
        Case .Item(3) And .Item(1) And .Item(0)
            Folder_Access = cNTFS.Return_Folder(Folder_Read_Write_List)
        Case .Item(2) And .Item(1) And .Item(0)
            Folder_Access = cNTFS.Return_Folder(Folder_Read_Execute_List)
        Case .Item(3) And .Item(0)
            Folder_Access = cNTFS.Return_Folder(Folder_Read_Write)
        Case .Item(2) And .Item(0)
            Folder_Access = cNTFS.Return_Folder(Folder_Read_Execute)
        Case .Item(1) And .Item(0)
            Folder_Access = cNTFS.Return_Folder(Folder_Read_List)
        Case .Item(0)
            Folder_Access = cNTFS.Return_Folder(Folder_Read)
        Case Else
            Folder_Access = -1
        End Select
    End With

End Function

Private Function Inheritence_Flags() As Long
'/* translate optionbox state to mask
    
    With cNTFS
        Select Case True
        Case optInherit(0).Value:       Inheritence_Flags = .Return_Inherit(Non_Propogate)
        Case optInherit(1).Value:       Inheritence_Flags = .Return_Inherit(Container_Inherit)
        Case optInherit(2).Value:       Inheritence_Flags = .Return_Inherit(Child_Container_Inherit)
        End Select
    End With
    
End Function

Private Function Access_Type() As Long
'/* translate optionbox state to mask

    With cNTFS
        Select Case True
        Case optAccess(0).Value:        Access_Type = .Return_Type(Access_Allowed)
        Case optAccess(1).Value:        Access_Type = .Return_Type(Access_Denied)
        End Select
    End With

End Function

Private Sub Select_File()
'/* get the file

On Error GoTo Handler

    With cdFile
        .DialogTitle = "Select a File"
        .CancelError = True
        .DefaultExt = ".txt"
        .InitDir = Left$(App.Path, 3)
        .ShowOpen
        txtNTFSPath.Text = .FileName
    End With

    If Len(txtNTFSPath.Text) > 0 Then
        stBar.SimpleText = "File: " + txtNTFSPath.Text + " selected.."
    End If
    
Handler:

End Sub

Private Sub NTFSDrive_Test()
'/* test ntfs availability
'/* create test directory

    If cNTFS.NTFS_Check Is Nothing Then
        MsgBox "None of your dives are formatted with the NTFS File System." & vbNewLine & _
        "NTFS related options have been Disabled.", vbExclamation, "No NTFS Drives Detected!"
        Disable_NTFS
        lblEFSStatus(0).Caption = "There are no NTFS Formatted Drives.."
        Exit Sub
    Else
        Dim vItem As Variant
        Dim sDrive As String
        For Each vItem In cNTFS.NTFS_Check
            sDrive = sDrive + CStr(vItem)
            m_NTFSDrive = vItem
        Next vItem
        lblEFSStatus(0).Caption = "NTFS Drives: " + sDrive
        cNTFS.Create_Directory m_NTFSDrive + "NMC\"
        txtNTFSPath.Text = m_NTFSDrive + "NMC\"
        dirNTFS.Path = m_NTFSDrive + "NMC\"
    End If

End Sub

Private Sub Disable_NTFS()

Dim oCtrl   As Control

On Error Resume Next

    For Each oCtrl In Controls
        If oCtrl.Tag = "NTRQ" Then
            oCtrl.Enabled = False
        End If
    Next oCtrl
    
On Error GoTo 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cNTFS = Nothing
End Sub
