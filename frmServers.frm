VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmServers 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Server Database"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDeleteAll 
      Caption         =   "&Delete All"
      Height          =   525
      Left            =   5670
      TabIndex        =   11
      Top             =   2670
      Width           =   3015
   End
   Begin MSComDlg.CommonDialog CDialog1 
      Left            =   960
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "&Import servers from UT favorites"
      Height          =   525
      Left            =   5670
      TabIndex        =   10
      Top             =   3300
      Width           =   3015
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   525
      Left            =   7230
      TabIndex        =   8
      Top             =   3900
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   525
      Left            =   5670
      TabIndex        =   7
      Top             =   3900
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add a server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   5670
      TabIndex        =   1
      Top             =   120
      Width           =   2985
      Begin VB.TextBox txtDesc 
         Height          =   585
         Left            =   630
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   270
         Width           =   2235
      End
      Begin VB.CommandButton cmdAddServer 
         Caption         =   "&Add"
         Height          =   375
         Left            =   150
         TabIndex        =   6
         Top             =   1650
         Width           =   2685
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   630
         TabIndex        =   5
         Top             =   1260
         Width           =   1095
      End
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   630
         TabIndex        =   4
         Top             =   930
         Width           =   2235
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desc :"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   330
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Port : "
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "IP :"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   990
         Width           =   240
      End
   End
   Begin VB.ListBox lstServers 
      Height          =   3570
      Left            =   90
      TabIndex        =   0
      Top             =   840
      Width           =   5445
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Picture         =   "frmServers.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "Right click on a server to edit or delete."
      Height          =   555
      Left            =   750
      TabIndex        =   9
      Top             =   240
      Width           =   3615
   End
   Begin VB.Menu mnuHidden 
      Caption         =   "HiddenMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUpdateDesc 
         Caption         =   "&Update Description"
      End
   End
End
Attribute VB_Name = "frmServers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ServerIP As String
Public ServerPort As String

Private bCancel As Boolean


Private Sub cmdAddServer_Click()
   If Trim(txtDesc) = "" Or Trim(txtIP) = "" Or Trim(txtPort) = "" Then Exit Sub
   lstServers.AddItem txtDesc & "\" & txtIP & "\" & txtPort
   txtIP = "": txtPort = ""
   lstServers.SetFocus
End Sub

Private Sub cmdCancel_Click()
   bCancel = True
   Unload Me
End Sub

Private Sub cmdDeleteAll_Click()
   If MsgBox("Are you sure you want to delete all servers?", vbQuestion + vbYesNo) = vbYes Then
      lstServers.Clear
   End If
End Sub

Private Sub cmdImport_Click()
   On Error GoTo ErrHandler
      
   Dim INILocation As String
   INILocation = GetSetting(App.Title, "Settings", "INILocation", "")
   
   If INILocation = "" Then
   
      MsgBox "You have to locate the file" & vbCrLf & vbCrLf & _
      "UnrealTournament.ini" & vbCrLf & vbCrLf & _
      "You will find it in the System subfolder of your UT folder.", vbInformation
      
      With CDialog1
         .CancelError = True
         .DialogTitle = "Find UnrealTournament.ini"
         .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNPathMustExist
         .Filter = "UnrealTournament.ini|UnrealTournament.ini"
         .ShowOpen
         INILocation = .FileName
         SaveSetting App.Title, "Settings", "INILocation", INILocation
      End With
      
   End If
      
   Screen.MousePointer = vbHourglass
      
   Dim INIData As String
   Open INILocation For Input As #1
   INIData = Input(LOF(1), 1)
   
   Dim Favs() As String
   Dim FavPos As Long
   Dim FavData As String
   Dim Slash1 As Long, Slash2 As Long, Slash3 As Long
   
   ReDim Preserve Favs(0)
   
   FavPos = InStr(1, INIData, "Favorites[", vbTextCompare)
   Do Until FavPos = 0
      FavData = Mid(INIData, FavPos, InStr(FavPos, INIData, Chr(13)) - FavPos)
      Slash1 = InStr(1, FavData, "\")
      If Slash1 = 0 Then Exit Do
      Slash2 = InStr(Slash1 + 1, FavData, "\")
      Slash3 = InStr(Slash2 + 1, FavData, "\")
      FavData = Left(FavData, Slash3 - 1)
      FavData = Mid(FavData, InStr(1, FavData, "=") + 1)
      
      ReDim Preserve Favs(UBound(Favs) + 1)
      Favs(UBound(Favs)) = FavData
      
      'Debug.Print FavData
      
      FavPos = InStr(FavPos + 1, INIData, "Favorites[", vbTextCompare)
      
   Loop
   
   Dim i As Long
   lstServers.Clear
   For i = 1 To UBound(Favs)
      lstServers.AddItem Favs(i)
   Next
      
CleanExit:
   
   Close #1
   Screen.MousePointer = vbDefault
   Exit Sub

ErrHandler:
   Select Case Err.Number
   Case 32755
      Resume CleanExit
   Case Else
      MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & Err.Source, vbCritical, "Error in frmServers.cmdImport.Click()"
      Resume CleanExit
   End Select
   
End Sub

Private Sub cmdOK_Click()
   If lstServers.ListIndex <> -1 Then
      ParseServerInfo lstServers.Text, ServerIP, ServerPort
   End If
   Unload Me
End Sub




Private Sub Form_Load()
   Dim AppPath As String
   If Right(App.Path, 1) = "\" Then
      AppPath = App.Path
   Else
      AppPath = App.Path & "\"
   End If
   
   On Error GoTo ErrHandler
   
   Dim sServer As String
   
   Open AppPath & "Servers.txt" For Input As #1
   Do Until EOF(1)
      Line Input #1, sServer
      lstServers.AddItem sServer
   Loop
   Close #1
   
   If lstServers.ListCount <> 0 Then lstServers.ListIndex = 0
   
CleanExit:

   Exit Sub

ErrHandler:

   Select Case Err
   Case 53
      Open AppPath & "Servers.txt" For Output As #1
      Resume
   Case Else
      MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & Err.Source, vbCritical, "Error in frmServers.Form_Load()"
      Resume CleanExit
   End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   If bCancel Then Exit Sub
   
   Me.Hide
      
   
   Dim AppPath As String
   If Right(App.Path, 1) = "\" Then
      AppPath = App.Path
   Else
      AppPath = App.Path & "\"
   End If
   
   Dim i As Long
   Open AppPath & "Servers.txt" For Output As #1
      For i = 0 To lstServers.ListCount - 1
         lstServers.ListIndex = i
         Print #1, lstServers.Text
      Next
   Close #1
   
End Sub

Private Sub ParseServerInfo(ByVal ServerInfo As String, lServerIP, lServerPort)
   Dim Slash1 As Long
   Dim Slash2 As Long
   Slash1 = InStr(1, ServerInfo, "\")
   Slash2 = InStr(Slash1 + 1, ServerInfo, "\")
   
   lServerIP = Mid(ServerInfo, Slash1 + 1, Slash2 - Slash1 - 1)
   lServerPort = Mid(ServerInfo, Slash2 + 1)
End Sub

Private Sub lstServers_DblClick()
   If lstServers.ListIndex = -1 Then Exit Sub
   ParseServerInfo lstServers.Text, ServerIP, ServerPort
   Unload Me
End Sub

Private Sub lstServers_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And lstServers.ListIndex <> -1 Then
   PopupMenu mnuHidden
End If
End Sub

Private Sub mnuDelete_Click()
   If lstServers.ListIndex <> -1 Then lstServers.RemoveItem lstServers.ListIndex
End Sub

Private Sub mnuEdit_Click()
   If lstServers.ListIndex = -1 Then Exit Sub
   Dim NewServerData As String
   NewServerData = InputBox("Enter the Descriotion,IP and port of the server," & vbCrLf & "in the format " & vbCrLf & vbCrLf & "Server Description\255.255.255.255\7777", "Edit Server", lstServers.List(lstServers.ListIndex))
   If NewServerData <> "" Then lstServers.List(lstServers.ListIndex) = NewServerData
End Sub

