VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NMC V1.5 - About"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7095
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
   ScaleHeight     =   4125
   ScaleWidth      =   7095
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Do It"
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
      Left            =   5160
      TabIndex        =   1
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox txtAbout 
      ForeColor       =   &H00404040&
      Height          =   3165
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmAbout.frx":0000
      Top             =   150
      Width           =   6825
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()

Dim sText   As String

    sText = " *** NT MASTER CLASS V1.6 *** " + vbNewLine + vbNewLine
    sText = sText + "/~ User and Group Management ~/" + vbNewLine + _
    "User_Create - create a new user" + vbNewLine + _
    "User_Delete - delete a user account" + vbNewLine + _
    "User_Exist - test a user account" + vbNewLine + _
    "User_Data - get user profile data" + vbNewLine + _
    "User_LoadProfile - create a user profile directory" + vbNewLine + _
    "User_UnloadProfile - unload user profile" + vbNewLine + _
    "User_LoadHive - load a user hive" + vbNewLine + _
    "User_UnloadHive - unload a user hive" + vbNewLine + _
    "Users_List - list local|global users" + vbNewLine + _
    "User_Impersonate - impersonate a user" + vbNewLine + _
    "User_Revert - revert to native token" + vbNewLine + _
    "User_RunAs - launch process as user" + vbNewLine + _
    "User_Name - logged in user" + vbNewLine + _
    "User_SetData - apply account changes" + vbNewLine + _
    "User_Password - change user password" + vbNewLine + _
    "Group_Create - create a user group" + vbNewLine + _
    "Group_Add - add user to group" + vbNewLine + _
    "Group_Remove - remove user from group" + vbNewLine + _
    "Group_Delete - delete a group" + vbNewLine + _
    "Groups_List - list built-in groups" + vbNewLine + _
    "Groups_List - list built-in groups" + vbNewLine + _
    "Get_Domain - get primary dc name" + vbNewLine + vbNewLine

    sText = sText + "/~ Service Management ~/" + vbNewLine + _
    "Service_Start - start a service" + vbNewLine + _
    "Service_Stop - stop a service" + vbNewLine + _
    "Service_Pause - pause a service" + vbNewLine + _
    "Service_Resume - start a service" + vbNewLine + _
    "Service_Query - query service state" + vbNewLine + _
    "Service_Change - change startup attr" + vbNewLine + _
    "Service_Add - add a new service" + vbNewLine + _
    "Service_Remove - delete a service" + vbNewLine + _
    "Service_Desc - change the service description" + vbNewLine + _
    "Service_Enumerate - list running services" + vbNewLine + vbNewLine

    sText = sText + "/~ Process Management ~/" + vbNewLine + _
    "Process_Enumerate - list running processes (kernal32)" + vbNewLine + _
    "Process_EnumG2 - server compliant enum (psapi)" + vbNewLine + _
    "Process_Exists - test existence (kernal32)" + vbNewLine + _
    "Process_ExistsG2 - test existence (psapi)" + vbNewLine + _
    "Process_GetClass - get process class" + vbNewLine + _
    "Process_SetClass - change process priority" + vbNewLine + _
    "Return_ProcessID - return prc id (kernal32)" + vbNewLine + _
    "Return_ProcessIDG2 - return prc id (psapi)" + vbNewLine + _
    "Process_Terminate - terminate a process" + vbNewLine + _
    "Thread_Enumerate - list a process threads" + vbNewLine + _
    "Thread_Suspend - suspend a thread" + vbNewLine + _
    "Thread_Resume - resume a thread" + vbNewLine + _
    "Thread_GetPriority - get thread priority" + vbNewLine + _
    "Thread_SetPriority - set thread priority" + vbNewLine + _
    "Thread_Terminate - kill a thread" + vbNewLine + _
    "Module_Enumerate - list a process modules" + vbNewLine + _
    "Module_EnumG2 - server compliant enum" + vbNewLine + vbNewLine

    sText = sText + "/~ Encrypted File System ~/" + vbNewLine + _
    "EFS_Status - file/folder state" + vbNewLine + _
    "EFS_Encrypt - encrypt file/folder" + vbNewLine + _
    "EFS_Decrypt - decrypt file/folder" + vbNewLine + _
    "EFS_Enable - enable EFS" + vbNewLine + _
    "EFS_Disable - disable EFS" + vbNewLine + _
    "Create_Directory - create a directory" + vbNewLine + _
    "File_Exists - test for path/file" + vbNewLine + vbNewLine
    
    sText = sText + "NTFS File System ~/" + vbNewLine + _
    "NTFS_Drive - test drive for ntfs" + vbNewLine + _
    "NTFS_Check - test all drives for ntfs" + vbNewLine + _
    "NTFS_Folder - modify object security" + vbNewLine + _
    "NTFS_Recursive - recurse permissions set" + vbNewLine + _
    "NTFS_Key - modify key security" + vbNewLine

    txtAbout.Text = sText
    
End Sub
