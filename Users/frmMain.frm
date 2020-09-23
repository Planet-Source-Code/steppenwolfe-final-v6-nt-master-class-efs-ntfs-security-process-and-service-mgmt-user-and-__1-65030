VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NMC - User Management"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9345
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
   ScaleHeight     =   6630
   ScaleWidth      =   9345
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3180
      TabIndex        =   24
      Top             =   1770
      Width           =   1305
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
      Left            =   3180
      TabIndex        =   23
      Top             =   1380
      Width           =   1305
   End
   Begin VB.TextBox txtUser 
      Height          =   315
      Index           =   3
      Left            =   4590
      TabIndex        =   18
      Text            =   "Test_Group"
      Top             =   5280
      Width           =   3765
   End
   Begin VB.TextBox txtUser 
      Height          =   315
      Index           =   2
      Left            =   4590
      TabIndex        =   17
      Top             =   4680
      Width           =   3765
   End
   Begin VB.TextBox txtUser 
      Height          =   315
      Index           =   1
      Left            =   4590
      TabIndex        =   16
      Text            =   "Test_Pass"
      Top             =   4080
      Width           =   3765
   End
   Begin VB.TextBox txtUser 
      Height          =   315
      Index           =   0
      Left            =   4590
      TabIndex        =   15
      Text            =   "Test_User"
      Top             =   3480
      Width           =   3765
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
      Height          =   525
      Index           =   6
      Left            =   180
      TabIndex        =   8
      Top             =   4050
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
      Height          =   525
      Index           =   5
      Left            =   180
      TabIndex        =   7
      Top             =   3420
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
      Height          =   525
      Index           =   4
      Left            =   180
      TabIndex        =   6
      Top             =   2790
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
      Height          =   525
      Index           =   3
      Left            =   180
      TabIndex        =   5
      Top             =   2160
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
      Height          =   525
      Index           =   2
      Left            =   180
      TabIndex        =   4
      Top             =   1530
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
      Height          =   525
      Index           =   1
      Left            =   180
      TabIndex        =   3
      Top             =   900
      Width           =   1695
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
      Height          =   525
      Index           =   0
      Left            =   180
      TabIndex        =   2
      Top             =   270
      Width           =   1695
   End
   Begin MSComctlLib.ListView lstUsers 
      Height          =   2895
      Left            =   4530
      TabIndex        =   0
      Top             =   210
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   5106
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   6285
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   609
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin VB.Label lblDisplay 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group Name:"
      Height          =   210
      Index           =   3
      Left            =   4590
      TabIndex        =   22
      Top             =   5070
      Width           =   945
   End
   Begin VB.Label lblDisplay 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Computer Name"
      Height          =   210
      Index           =   2
      Left            =   4590
      TabIndex        =   21
      Top             =   4470
      Width           =   1140
   End
   Begin VB.Label lblDisplay 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   210
      Index           =   1
      Left            =   4590
      TabIndex        =   20
      Top             =   3870
      Width           =   795
   End
   Begin VB.Label lblDisplay 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      Height          =   210
      Index           =   0
      Left            =   4590
      TabIndex        =   19
      Top             =   3270
      Width           =   840
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   210
      Index           =   5
      Left            =   180
      TabIndex        =   14
      Top             =   5880
      Width           =   45
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   210
      Index           =   4
      Left            =   180
      TabIndex        =   13
      Top             =   5640
      Width           =   45
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   210
      Index           =   3
      Left            =   180
      TabIndex        =   12
      Top             =   5400
      Width           =   45
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   210
      Index           =   2
      Left            =   180
      TabIndex        =   11
      Top             =   5160
      Width           =   45
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   210
      Index           =   1
      Left            =   180
      TabIndex        =   10
      Top             =   4920
      Width           =   45
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   210
      Index           =   0
      Left            =   180
      TabIndex        =   9
      Top             =   4680
      Width           =   45
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cUser    As clsUsrMgnt
Attribute cUser.VB_VarHelpID = -1
Private m_cItem             As ListItem

Private Sub cmdUsrControls_Click(Index As Integer)

    With cUser
        Select Case Index
        '/* create user
        Case 0
            If .User_Create(txtUser(2).Text, txtUser(0).Text, _
                txtUser(2).Text, App.Path, "Test User", App.Path + "\Scripts") Then
            End If
            Get_Users True
            
        '/* user exists
        Case 1
            .User_Exist txtUser(2).Text, txtUser(0).Text
            
        '/* create a group
        Case 2
            If .Group_Create(txtUser(2).Text, txtUser(3).Text, "Test Comment") Then
                Get_Groups
            End If
        
        '/* add to group
        Case 3
            .Group_Add txtUser(2).Text, txtUser(0).Text, txtUser(3).Text
            
        '/* user data
        Case 4
            On Error Resume Next
            Dim cTemp As Collection
            Set cTemp = New Collection
            Set cTemp = cUser.User_Data(txtUser(2).Text, txtUser(0).Text)
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
            Get_Groups
            
        '/* list groups
        Case 7
            Get_Groups
            
        '/* list users
        Case 8
            Get_Users True
        End Select
    End With
    
End Sub

Private Sub cUser_eNComplete(ByVal sTask As String)
'/* task completed
    stBar.SimpleText = sTask
End Sub

Private Sub cUser_eNErrorCond(ByVal sRoutine As String, ByVal sError As String)
'/* error condition
    MsgBox "The Routine: " + sRoutine + " could not complete.. Error: " + sError
End Sub

Private Sub Form_Load()

    Set cUser = New clsUsrMgnt
    txtUser(2).Text = cUser.Computer_Name
    Get_Groups

End Sub

Private Sub Get_Users(ByVal bLocal As Boolean)

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
    
    For Each vItem In cUser.Users_List(txtUser(2).Text, "", bLocal)
        If Len(vItem) = 0 Then GoTo skip
        Set m_cItem = lstUsers.ListItems.Add(Text:=vItem)
skip:
    Next vItem

On Error GoTo 0

End Sub

Private Sub Get_Groups()

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
    
    For Each vItem In cUser.Groups_List
        sUData = Split(CStr(vItem), Chr$(31))
        If Len(sUData(0)) = 0 Then GoTo skip
        Set m_cItem = lstUsers.ListItems.Add(Text:=sUData(0))
        m_cItem.SubItems(1) = sUData(1)
        m_cItem.SubItems(2) = sUData(2)
skip:
    Next vItem
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cUser = Nothing
End Sub
