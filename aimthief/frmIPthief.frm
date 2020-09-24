VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmIPthief 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AIM IP Thief"
   ClientHeight    =   1800
   ClientLeft      =   3615
   ClientTop       =   3330
   ClientWidth     =   4065
   Icon            =   "frmIPthief.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4065
   Begin MSWinsockLib.Winsock Aim2 
      Left            =   4920
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.OptionButton FLSend 
      BackColor       =   &H00E0E0E0&
      Caption         =   "File Send"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.OptionButton DRConnect 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Direct Connect"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.ListBox lstIP 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin MSWinsockLib.Winsock AIM 
      Left            =   4440
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   120
      Picture         =   "frmIPthief.frx":038A
      Top             =   1440
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   2040
      Picture         =   "frmIPthief.frx":13CC
      Top             =   1440
      Width           =   240
   End
End
Attribute VB_Name = "frmIPthief"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//////// Written By Violent        Violence 1999  -   2002 ///////////

Private Sub DRConnect_Click()
AIM.Close
If DRConnect.Value = True Then: AIM.LocalPort = 4443: AIM.Listen
End Sub

Private Sub FLSend_Click()
Aim2.Close
If FLSend.Value = True Then: Aim2.LocalPort = 5190: Aim2.Listen
End Sub


Private Sub AIM_ConnectionRequest(ByVal requestID As Long)
lstIP.AddItem "Time: " & Time & " " & AIM.RemoteHostIP & " Direct Connect"
End Sub


Private Sub AIM2_ConnectionRequest(ByVal requestID As Long)
lstIP.AddItem "Time: " & Time & " " & Aim2.RemoteHostIP & " File Send"
End Sub

