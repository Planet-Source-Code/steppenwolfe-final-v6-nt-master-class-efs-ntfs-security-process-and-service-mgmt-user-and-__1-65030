VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NMC - Process Management"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10230
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
   ScaleHeight     =   7245
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdControls 
      Caption         =   "Thread Resume"
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
      Left            =   3600
      TabIndex        =   8
      Top             =   6030
      Width           =   1635
   End
   Begin VB.CommandButton cmdControls 
      Caption         =   "Thread Suspend"
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
      Left            =   1860
      TabIndex        =   7
      Top             =   6030
      Width           =   1635
   End
   Begin VB.CommandButton cmdControls 
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
      Index           =   4
      Left            =   150
      TabIndex        =   6
      Top             =   6030
      Width           =   1635
   End
   Begin VB.CommandButton cmdControls 
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
      Index           =   3
      Left            =   5340
      TabIndex        =   5
      Top             =   5520
      Width           =   1635
   End
   Begin VB.CommandButton cmdControls 
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
      Left            =   3600
      TabIndex        =   3
      Top             =   5520
      Width           =   1635
   End
   Begin VB.CommandButton cmdControls 
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
      Left            =   1860
      TabIndex        =   2
      Top             =   5520
      Width           =   1635
   End
   Begin VB.CommandButton cmdControls 
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
      Left            =   150
      TabIndex        =   1
      Top             =   5520
      Width           =   1635
   End
   Begin MSComctlLib.ListView lstProcess 
      Height          =   5115
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   9022
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
      TabIndex        =   4
      Top             =   6900
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   609
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cProc    As clsProcess
Attribute cProc.VB_VarHelpID = -1
Private m_cItem             As ListItem

Private Sub cmdControls_Click(Index As Integer)

    With cProc
        Select Case Index
        '/* process list
        Case 0
            Get_Processes
            
        '/* process exists
        Case 1
            .Process_Exists "alg.exe", True
            
        '/* kill process
        Case 2
            .Process_Terminate lstProcess.SelectedItem
        
        '/* module list
        Case 3
            If Not .Process_Exists(lstProcess.SelectedItem.Text) Then
                MsgBox "Please select a valid process before continuing!", vbExclamation, "No Process Name!"
                Exit Sub
            End If
            Get_Modules lstProcess.SelectedItem.Text
        
        '/* thread list
        Case 4
            If Not .Process_Exists(lstProcess.SelectedItem.Text) Then
                MsgBox "Please select a valid process before continuing!", vbExclamation, "No Process Name!"
                Exit Sub
            End If
            Get_Threads lstProcess.SelectedItem.Text
        
        '/* suspend thread
        Case 5
            If Not IsNumeric(lstProcess.SelectedItem.Text) Then
                MsgBox "Please select a valid Thread ID before continuing!", vbExclamation, "Invalid Input"
                Exit Sub
            End If
            .Thread_Suspend CLng(lstProcess.SelectedItem.Text), ""
            
        '/* resume thread
        Case 6
            If Not IsNumeric(lstProcess.SelectedItem.Text) Then
                MsgBox "Please select a valid Thread ID before continuing!", vbExclamation, "Invalid Input"
                Exit Sub
            End If
            .Thread_Resume CLng(lstProcess.SelectedItem.Text), ""
        End Select
    End With

End Sub

Private Sub cProc_eNComplete(ByVal sTask As String)
    stBar.SimpleText = sTask
End Sub

Private Sub cProc_eNErrorCond(ByVal sRoutine As String, _
                              ByVal sError As String)

    MsgBox "The Routine: " + sRoutine + " could not complete.. Error: " + sError

End Sub

Private Sub Form_Load()

    Set cProc = New clsProcess
    Get_Processes

End Sub

Private Sub Get_Processes()

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
    
    Set cTemp = cProc.Process_Enumerate
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

Private Sub Get_Modules(ByVal sProcess As String)

Dim vItem       As Variant
Dim cTemp       As Collection
Dim sUData()    As String

On Error Resume Next

    '/* check for mods first
    Set cTemp = cProc.Module_Enumerate(sProcess)
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

Private Sub Get_Threads(ByVal sProcess As String)

Dim vItem       As Variant
Dim cTemp       As Collection
Dim sUData()    As String

On Error Resume Next

    '/* check for mods first
    Set cTemp = cProc.Thread_Enumerate(sProcess)
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
