VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "UTServer Query Tool v1.0"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7485
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   5715
      Left            =   0
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   5655
      ScaleWidth      =   7455
      TabIndex        =   0
      Top             =   0
      Width           =   7515
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theok21@internet.gr"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   2
         Left            =   1620
         TabIndex        =   3
         Top             =   2670
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "20/05/2001"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   5250
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Programmed in MS Visual Basic 6.0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   5010
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Unload Me
End Sub

Private Sub Label1_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Picture1_Click()
   Unload Me
End Sub
