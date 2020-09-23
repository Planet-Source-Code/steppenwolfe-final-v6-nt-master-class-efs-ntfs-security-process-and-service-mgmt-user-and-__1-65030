VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NMC - Service Management"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10755
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
   ScaleHeight     =   6165
   ScaleWidth      =   10755
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optStartup 
      Caption         =   "Disable"
      Height          =   195
      Index           =   2
      Left            =   9210
      TabIndex        =   10
      Top             =   5580
      Width           =   855
   End
   Begin MSComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   9
      Top             =   5820
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   609
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin VB.OptionButton optStartup 
      Caption         =   "Manual"
      Height          =   195
      Index           =   1
      Left            =   9210
      TabIndex        =   8
      Top             =   5340
      Value           =   -1  'True
      Width           =   885
   End
   Begin VB.OptionButton optStartup 
      Caption         =   "Auto"
      Height          =   195
      Index           =   0
      Left            =   9210
      TabIndex        =   7
      Top             =   5100
      Width           =   825
   End
   Begin VB.CommandButton cmdControls 
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
      TabIndex        =   6
      Top             =   5130
      Width           =   1365
   End
   Begin VB.CommandButton cmdControls 
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
      TabIndex        =   5
      Top             =   5130
      Width           =   1365
   End
   Begin VB.CommandButton cmdControls 
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
      TabIndex        =   4
      Top             =   5130
      Width           =   1365
   End
   Begin VB.CommandButton cmdControls 
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
      TabIndex        =   3
      Top             =   5130
      Width           =   1365
   End
   Begin VB.CommandButton cmdControls 
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
      TabIndex        =   2
      Top             =   5130
      Width           =   1365
   End
   Begin VB.CommandButton cmdControls 
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
      TabIndex        =   1
      Top             =   5130
      Width           =   1365
   End
   Begin MSComctlLib.ListView lstServices 
      Height          =   4785
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   8440
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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private WithEvents cServ As clsServices
Attribute cServ.VB_VarHelpID = -1

Private Sub cmdControls_Click(Index As Integer)

Dim lItem   As ListItem

On Error GoTo Handler

    With cServ
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
        cmdControls_Click 0
        stBar.SimpleText = sOldText
    End If
    
Handler:
    On Error GoTo 0

End Sub

Private Sub cServ_eNComplete(ByVal sTask As String)
    stBar.SimpleText = sTask
End Sub

Private Sub cServ_eNErrorCond(ByVal sRoutine As String, ByVal sError As String)
    MsgBox "The Routine: " + sRoutine + " could not complete.. Error: " + sError
End Sub

Private Sub Form_Load()

    Set cServ = New clsServices
    With lstServices
        .View = lvwReport
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .AllowColumnReorder = True
        .ColumnHeaders.Add 1, , "Display Name", .Width / 3
        .ColumnHeaders.Add 2, , "Service Name", .Width / 3
        .ColumnHeaders.Add 3, , "Service Status", (.Width / 3) - 100
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cServ = Nothing
End Sub
