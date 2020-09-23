VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NMC - Encrypted File System"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10605
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   10605
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEFSPath 
      Height          =   315
      Left            =   7620
      TabIndex        =   15
      Tag             =   "NTRQ"
      Top             =   5520
      Width           =   2775
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
      Left            =   5730
      TabIndex        =   13
      Tag             =   "NTRQ"
      Top             =   5460
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
      Left            =   3900
      TabIndex        =   12
      Tag             =   "NTRQ"
      Top             =   5460
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
      Left            =   2070
      TabIndex        =   11
      Tag             =   "NTRQ"
      Top             =   5460
      Width           =   1695
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
      Left            =   240
      TabIndex        =   10
      Tag             =   "NTRQ"
      Top             =   5460
      Width           =   1695
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
      Left            =   3090
      TabIndex        =   9
      Tag             =   "NTRQ"
      Top             =   4020
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
      Left            =   3090
      TabIndex        =   8
      Tag             =   "NTRQ"
      Top             =   2730
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
      Left            =   3090
      TabIndex        =   7
      Tag             =   "NTRQ"
      Top             =   1470
      Width           =   2115
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
      Left            =   3090
      TabIndex        =   2
      Tag             =   "NTRQ"
      Top             =   210
      Width           =   2115
   End
   Begin VB.TextBox txtEFS 
      Height          =   4335
      Left            =   5460
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Tag             =   "NTRQ"
      Top             =   150
      Width           =   4995
   End
   Begin MSComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   6045
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   609
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdFile 
      Left            =   10020
      Top             =   4590
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   7650
      TabIndex        =   16
      Top             =   5310
      Width           =   405
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
      Left            =   5460
      TabIndex        =   14
      Top             =   4530
      Width           =   1020
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
      Left            =   270
      TabIndex        =   6
      Top             =   4050
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
      Left            =   270
      TabIndex        =   5
      Top             =   2790
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
      Left            =   270
      TabIndex        =   4
      Top             =   1470
      Width           =   2805
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
      Left            =   270
      TabIndex        =   3
      Top             =   210
      Width           =   2805
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cEFS     As clsEFS
Attribute cEFS.VB_VarHelpID = -1
Private m_NTFSDrive         As String


Private Sub cEFS_eNComplete(ByVal sTask As String)
'/* task completed
    stBar.SimpleText = sTask
End Sub

Private Sub cEFS_eNErrorCond(ByVal sRoutine As String, ByVal sError As String)
'/* error condition
    MsgBox "The Routine: " + sRoutine + " could not complete.. Error: " + sError
End Sub

Private Sub cmdEFSDemo_Click(Index As Integer)
'/* demo

Dim sPath       As String

    sPath = m_NTFSDrive + "NMC\sample.txt"
    
    With cEFS
        Select Case Index
        '/* setup
        Case 0
            CopyFile App.Path + "\sample.txt", sPath, 0
            txtEFS.Text = File_Data(sPath)
            cmdEFSDemo(1).Enabled = True
            lblEFS(0).ForeColor = &H404040
            lblEFS(1).ForeColor = &HC00000
        
        '/* encrypt
        Case 1
            txtEFS.Text = ""
            If .EFS_Encrypt(sPath) Then
                Debug.Print "encryption success!"
            End If
            txtEFS.Text = File_Data(sPath)
            cmdEFSDemo(2).Enabled = True
            lblEFS(1).ForeColor = &H404040
            lblEFS(2).ForeColor = &HC00000
            
        '/* impersonate and view
        Case 2
            If .User_Impersonate("EFS_Test", "password", .Local_Name) Then
                Debug.Print "impersonate success!"
                txtEFS.Text = ""
            Else
                MsgBox "User: EFS_Test Password: password does not exist!" & vbNewLine & _
                "Please create the account manually, or run the test from the main project.", _
                vbInformation, "Account does not exist!"
                Exit Sub
            End If
            txtEFS.Text = File_Data(sPath)
            cmdEFSDemo(3).Enabled = True
            lblEFS(2).ForeColor = &H404040
            lblEFS(3).ForeColor = &HC00000
        
        '/* revert token and reset
        Case 3
            If .User_Revert Then
                Debug.Print "account reverted!"
            End If
            .EFS_Decrypt sPath
            txtEFS.Text = File_Data(sPath)
        End Select
    End With
    
End Sub

Private Function File_Data(ByVal sPath As String) As String

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
    Debug.Print Err.Description + " Err# " + CStr(Err.Number)

End Function

Private Sub Form_Load()

    Set cEFS = New clsEFS
    NTFSDrive_Test
    
End Sub

Private Sub cmdEFSControls_Click(Index As Integer)

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
        With cEFS
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
        With cEFS
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

Private Sub NTFSDrive_Test()
'/* test ntfs availability
'/* create test directory

    If cEFS.NTFS_Check Is Nothing Then
        MsgBox "None of your dives are formatted with the NTFS File System." & vbNewLine & _
        "NTFS related options have been Disabled.", vbExclamation, "No NTFS Drives Detected!"
        Disable_NTFS
        lblEFSStatus(0).Caption = "There are no NTFS Formatted Drives.."
        Exit Sub
    Else
        Dim vItem As Variant
        Dim sDrive As String
        For Each vItem In cEFS.NTFS_Check
            sDrive = sDrive + CStr(vItem)
            m_NTFSDrive = vItem
        Next vItem
        lblEFSStatus(0).Caption = "NTFS Drives: " + sDrive
        cEFS.Create_Directory m_NTFSDrive + "NMC\"
        txtEFSPath.Text = m_NTFSDrive + "NMC\"
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
    Set cEFS = Nothing
End Sub
