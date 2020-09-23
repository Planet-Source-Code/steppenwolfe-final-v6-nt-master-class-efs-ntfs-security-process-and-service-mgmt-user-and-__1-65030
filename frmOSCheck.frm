VERSION 5.00
Begin VB.Form frmOSCheck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NMC V1.5 - OS Check"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5550
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
   ScaleHeight     =   2925
   ScaleWidth      =   5550
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Caption         =   "OK"
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
      Left            =   4200
      TabIndex        =   0
      Top             =   2160
      Width           =   1125
   End
   Begin VB.Label lblData 
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
      Height          =   1320
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   750
      Width           =   5190
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Header"
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
      Left            =   120
      TabIndex        =   2
      Top             =   450
      Width           =   945
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Header"
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
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmOSCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_sOSType As String

Public Property Let p_OSType(ByVal PropVal As String)
    m_sOSType = PropVal
End Property

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    lblData(0).Caption = "Some Features will Not be available!"
    lblData(1).Caption = "Your OS: " + m_sOSType
    lblData(2).Caption = "Some features of this demonstration have been " & _
        "disabled because your operating system does not support them. " & _
        "NFS, Registry and File System NTFS demonstrations " & _
        "have been disabled. The User/Group, Process and Service Management " & _
        "demonstrations are [in part], backwards compatable."

End Sub
