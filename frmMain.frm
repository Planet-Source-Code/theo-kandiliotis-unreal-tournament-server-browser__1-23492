VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmServerQuery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UT Server Query Tool"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9945
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   9945
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CDialog1 
      Left            =   2580
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   90
      Top             =   180
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   4125
      Left            =   5520
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2490
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   7276
      _Version        =   393216
      BackColorBkg    =   -2147483633
      AllowBigSelection=   0   'False
      TextStyle       =   3
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
      BorderStyle     =   0
   End
   Begin VB.TextBox txtInfo 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4485
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2130
      Width           =   5295
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5250
      Top             =   -60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Frame Frame1 
      Height          =   2325
      Left            =   5520
      TabIndex        =   9
      Top             =   60
      Width           =   4365
      Begin VB.CheckBox chkNotify 
         Caption         =   "&Notify me when someone connects to the server"
         Height          =   285
         Left            =   180
         TabIndex        =   14
         Top             =   1410
         Width           =   3855
      End
      Begin VB.CommandButton cmdPlay 
         Cancel          =   -1  'True
         Caption         =   "&Play"
         CausesValidation=   0   'False
         Height          =   405
         Left            =   1530
         TabIndex        =   5
         Top             =   1770
         Width           =   1305
      End
      Begin VB.CheckBox chkAutoUpdate 
         Caption         =   "&Auto update every 10 seconds"
         Height          =   285
         Left            =   180
         TabIndex        =   3
         Top             =   1080
         Width           =   3525
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   3540
         TabIndex        =   2
         ToolTipText     =   "Click here to browse your favorite servers"
         Top             =   240
         Width           =   645
      End
      Begin VB.TextBox txtServerIP 
         Height          =   285
         Left            =   1110
         TabIndex        =   0
         Text            =   "194.134.233.76"
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "&Query"
         Default         =   -1  'True
         Height          =   405
         Left            =   180
         TabIndex        =   4
         Top             =   1770
         Width           =   1305
      End
      Begin VB.TextBox txtServerPort 
         Height          =   285
         Left            =   1110
         TabIndex        =   1
         Text            =   "7778"
         Top             =   630
         Width           =   2055
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         CausesValidation=   0   'False
         Height          =   405
         Left            =   2880
         TabIndex        =   6
         Top             =   1770
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Server IP:"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   300
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Server Port:"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   690
         Width           =   840
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   270
      MouseIcon       =   "frmMain.frx":08CA
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":0BD4
      ScaleHeight     =   1935
      ScaleWidth      =   4935
      TabIndex        =   12
      Top             =   90
      Width           =   4935
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Server Query Tool"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1230
         TabIndex        =   13
         Top             =   1530
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmServerQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const TIMEOUT_PERIOD = 3

Private Enum enDataState
   dsNotQuerying = 0
   dsQuerying = 1
End Enum

Private enumQueryState As enDataState

Private bFinalPacketIsIn As Boolean
Private sFinalPacket As String


Private Packets() As String
Private GridSortCol As Long
Private ColSortAsc(4) As Boolean

Private bEmptyServer As Boolean
Private PrevServer As String

Private PingTimer_Start As Single
Private PingTimer_End As Single
Private PingTime As Long

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Const SND_ASYNC = &H1
Private Const SND_LOOP = &H8
Private Const SND_MEMORY = &H4
Private Const SND_NODEFAULT = &H2
Private Const SND_APPLICATION = &H80
Private Const SND_NOSTOP = &H10
Private Const SND_NOWAIT = &H2000
Private Const SND_PURGE = &H40
Private Const SND_RESERVED = &HFF000000
Private Const SND_RESOURCE = &H40004
Private Const SND_SYNC = &H0
Private Const SND_TYPE_MASK = &H170007
Private Const SND_VALID = &H1F
Private Const SND_VALIDFLAGS = &H17201F



Private Sub chkAutoUpdate_Click()
   If enumQueryState = dsQuerying Then Exit Sub
   tmrUpdate.Enabled = (chkAutoUpdate = vbChecked)
   If (chkAutoUpdate = vbChecked) Then
      cmdQuery.Enabled = False
   Else
      cmdQuery.Enabled = True
   End If
   
End Sub

Private Sub chkNotify_Click()
   If chkNotify = vbChecked Then bEmptyServer = True
End Sub

Private Sub cmdBrowse_Click()
   frmServers.Show vbModal
   If frmServers.ServerIP <> "" Then
      txtServerIP = frmServers.ServerIP
      txtServerPort = frmServers.ServerPort
   End If
   Unload frmServers: Set frmServers = Nothing
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdPlay_Click()

   If txtServerIP = "" Or txtServerPort = "" Then Exit Sub
   
   Dim EXELocation As String
   EXELocation = GetSetting(App.Title, "Settings", "EXELocation", "")
   
   If EXELocation = "" Then
   
      MsgBox "You have to locate the file" & vbCrLf & vbCrLf & _
      "UnrealTournament.exe" & vbCrLf & vbCrLf & _
      "You will find it in the System subfolder of your UT folder.", vbInformation
      
      With CDialog1
         .CancelError = True
         .DialogTitle = "Find UnrealTournament.exe"
         .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNPathMustExist
         .Filter = "UnrealTournament.exe|UnrealTournament.exe"
         .ShowOpen
         EXELocation = .FileName
         SaveSetting App.Title, "Settings", "EXELocation", EXELocation
      End With
      
   End If

   Me.WindowState = vbMinimized
   
   Shell EXELocation & " " & txtServerIP & ":" & txtServerPort

End Sub

'Private Sub Ping()
'
'   enumQueryState = dsQuerying
'   bPinging = True
'   PingTimer_Start = Timer
'   Winsock1.SendData "\echo\"
'   Dim PingStartTimer As Single, TimedOut As Boolean
'   PingStartTimer = Timer
'
'   Do Until bPinging = False
'      DoEvents
'      If Timer > PingStartTimer + TIMEOUT_PERIOD Then TimedOut = True
'   Loop
'
'   If TimedOut Then
'      txtInfo = txtInfo & vbCrLf & "PING            : 9999"
'   Else
'      txtInfo = txtInfo & vbCrLf & "PING            : " & Int((PingTimer_End - PingTimer_Start) * 1000)
'   End If
'
'   enumQueryState = dsNotQuerying
'
'
'End Sub

Private Sub cmdQuery_Click()

   Static bCanceled As Boolean

   If cmdQuery.Caption = "&Query" Then
      
      cmdQuery.Caption = "&Cancel"

      On Error GoTo ErrHandler
   
      If enumQueryState = dsQuerying Then Exit Sub
   
      enumQueryState = dsQuerying
      txtInfo = ""
      Flex.Cols = 4
      Flex.Rows = 2
      Flex.TextMatrix(1, 0) = "": Flex.TextMatrix(1, 1) = "": Flex.TextMatrix(1, 2) = "": Flex.TextMatrix(1, 3) = ""
      ReDim Packets(0)
      bFinalPacketIsIn = False
      sFinalPacket = ""
      
      If Winsock1.State = 1 Then Winsock1.Close
      
      If txtServerIP & ":" & txtServerPort <> PrevServer Then bEmptyServer = True
      PrevServer = txtServerIP & ":" & txtServerPort
      
      Dim StartTimer As Single
      With Winsock1
         .RemoteHost = txtServerIP
         .RemotePort = txtServerPort
         enumQueryState = dsQuerying
         
         Screen.MousePointer = vbHourglass
         PingTimer_Start = Timer
         PingTimer_End = 0
         PingTime = 0
         .SendData "\status\"
      End With
      
      
      Dim TimedOut As Boolean
      
      StartTimer = Timer
      Do
         DoEvents
         If Timer > StartTimer + TIMEOUT_PERIOD Then TimedOut = True
      Loop Until (enumQueryState = dsNotQuerying) Or TimedOut
      
      If Not bCanceled Then
      
         If Not TimedOut Then
            SortPackets
            ProcessPackets
            SortGrid GridSortCol, True
            txtInfo = txtInfo & vbCrLf & "PING            : " & Int((PingTimer_End - PingTimer_Start) * 1000)
         Else
            txtInfo = "Query Timed Out." & vbCrLf & "The server is down or your ping" & vbCrLf & "is extremely high (>3000) ."
         End If
      Else
         bCanceled = False
      End If
      
   
Else
   
   bCanceled = True
   cmdQuery.Caption = "&Query"
   enumQueryState = dsNotQuerying

End If

   
CleanExit:
   
   cmdQuery.Caption = "&Query"
   enumQueryState = dsNotQuerying
   Screen.MousePointer = vbNormal
   
   Exit Sub

ErrHandler:

   If Err = 126 And Err.Source = "Winsock" Then
      Resume
   Else
      MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & Err.Source, vbCritical, "Error in frmServerQuery.cmdQuery.Click()"
      Resume CleanExit
   End If
   
End Sub

Private Sub Notify()
   Dim AppPath As String
   If Right(App.Path, 1) = "\" Then
      AppPath = App.Path
   Else
      AppPath = App.Path & "\"
   End If
   
   sndPlaySound AppPath & "m2altfire.wav", SND_ASYNC
   
End Sub

Private Sub ProcessPackets()
   Dim i As Long
   Dim RawData As String
   
   For i = 1 To UBound(Packets)
      RawData = RawData & Packets(i)
   Next
   
   If GetData(RawData, "gamename") <> "ut" Then
      txtInfo = "The server is hosting a game but it's not " & vbCrLf & _
             "an Unreal Tournament game !"
      Exit Sub
   End If
   
   txtInfo = _
   "Game ver.       : " & GetData(RawData, "gamever") & vbCrLf & _
   "Min. comp. ver. : " & GetData(RawData, "minnetver") & vbCrLf & _
   "Server name     : " & GetData(RawData, "hostname") & vbCrLf & _
   "Server port     : " & GetData(RawData, "hostport") & vbCrLf & _
   "Current map     : " & GetData(RawData, "maptitle") & " (" & GetData(RawData, "mapname") & ")" & vbCrLf & _
   "Game type       : " & GetData(RawData, "gametype") & vbCrLf & _
   "Players         : " & GetData(RawData, "numplayers") & "/" & GetData(RawData, "maxplayers") & vbCrLf & _
   "Game mode       : " & GetData(RawData, "gamemode") & vbCrLf & _
   "ngWorldStats    : " & IIf(GetData(RawData, "worldlog"), "Active", "InActive") & vbCrLf & _
   "Password req.   : " & GetData(RawData, "password") & vbCrLf & _
   "Time limit      : " & GetData(RawData, "timelimit") & vbCrLf & _
   "Frag limit      : " & GetData(RawData, "fraglimit") & vbCrLf & _
   "Min. players    : " & GetData(RawData, "minplayers") & vbCrLf & _
   "Change levels   : " & GetData(RawData, "changelevels") & vbCrLf & _
   "Tournament mode : " & GetData(RawData, "tournament") & vbCrLf & _
   "Game style      : " & GetData(RawData, "gamestyle") & vbCrLf & _
   "Admin name      : " & GetData(RawData, "AdminName") & vbCrLf & _
   "Admin email     : " & GetData(RawData, "AdminEMail")

   Dim PlayerCount As String
   PlayerCount = GetData(RawData, "numplayers")
   
   If bEmptyServer And (PlayerCount <> 0) And (chkNotify = vbChecked) Then Notify
   
   bEmptyServer = (PlayerCount = 0)
   
   For i = 0 To PlayerCount - 1
      If i = 0 Then
         Flex.TextMatrix(1, 0) = GetData(RawData, "player_" & i)
         Flex.TextMatrix(1, 1) = GetData(RawData, "frags_" & i)
         Flex.TextMatrix(1, 2) = GetData(RawData, "ping_" & i)
         Flex.TextMatrix(1, 3) = GetData(RawData, "team_" & i)
      Else
         Flex.AddItem GetData(RawData, "player_" & i) & vbTab & GetData(RawData, "frags_" & i) & vbTab & GetData(RawData, "ping_" & i) & vbTab & GetData(RawData, "team_" & i)
      End If
   Next


   

End Sub

Private Function GetData(ByVal RawData As String, ByVal Column As String) As String
   Dim DataPos As Long
   
   DataPos = InStr(1, RawData, Column) + Len(Column) + 1
   
   GetData = Mid(RawData, DataPos, InStr(DataPos, RawData, "\") - DataPos)
   
End Function

Private Sub SortPackets()

   Dim loop1 As Long, loop2 As Long, Y As Long, temp As String
   For loop1 = UBound(Packets) To 1 Step -1
      For loop2 = 2 To loop1
         
         Dim QueryNumPos As Long, QueryNum1 As String, QueryNum2 As String, s As String
         
         'get querynum for 1st packet
         s = Packets(loop2 - 1)
         QueryNumPos = InStr(InStr(1, s, "\queryid\"), s, ".")
         If InStr(QueryNumPos, s, "\") Then
            QueryNum1 = Mid(s, QueryNumPos + 1, InStr(QueryNumPos, s, "\") - QueryNumPos - 1)
         Else
            QueryNum1 = Mid(s, QueryNumPos + 1)
         End If
         
         If InStr(1, QueryNum1, "\") Then Stop
         
         'get querynum for 2nd packet
         s = Packets(loop2)
         QueryNumPos = InStr(InStr(1, s, "\queryid\"), s, ".")
         If InStr(QueryNumPos, s, "\") Then
            QueryNum2 = Mid(s, QueryNumPos + 1, InStr(QueryNumPos, s, "\") - QueryNumPos - 1)
         Else
            QueryNum2 = Mid(s, QueryNumPos + 1)
         End If
         
         If CLng(QueryNum1) > CLng(QueryNum2) Then
            temp = Packets(loop2 - 1)
            Packets(loop2 - 1) = Packets(loop2)
            Packets(loop2) = temp
         End If
         
         
      Next
   Next

End Sub


Private Sub Flex_Click()
   If Flex.MouseRow = 0 Then SortGrid Flex.MouseCol
End Sub

Private Sub SortGrid(ByVal Col As Long, Optional ByVal NoToggle = False)
   Flex.Col = Col
   Flex.ColSel = Col
   
   GridSortCol = Col
   
   
   
   Select Case Col
   Case 0
      If ColSortAsc(Col) Then
         Flex.Sort = flexSortStringAscending
      Else
         Flex.Sort = flexSortStringDescending
      End If
   Case 1, 2, 3
      If ColSortAsc(Col) Then
         Flex.Sort = flexSortNumericAscending
      Else
         Flex.Sort = flexSortNumericDescending
      End If
      
   End Select
   
   If NoToggle = False Then ColSortAsc(Col) = Not ColSortAsc(Col)
End Sub

Private Sub Form_Load()

   txtServerIP = GetSetting(App.Title, "Settings", "txtServerIp", "")
   txtServerPort = GetSetting(App.Title, "Settings", "txtServerPort", "")
   chkAutoUpdate = GetSetting(App.Title, "Settings", "chkAutoUpdate", vbUnchecked)
   chkNotify = GetSetting(App.Title, "Settings", "chkNotify", vbUnchecked)

   Dim i As Long

   GridSortCol = 1
   
   With Flex
      .Cols = 4
      .Rows = 2
      .FixedRows = 1
      .FixedCols = 0
      
      Flex.TextMatrix(0, 0) = "Nick"
      Flex.TextMatrix(0, 1) = "Frags"
      Flex.TextMatrix(0, 2) = "Ping"
      Flex.TextMatrix(0, 3) = "Team"
      
      Flex.Row = 0
      For i = 0 To 3
         Flex.Col = i
         Flex.CellFontBold = True
         Flex.CellTextStyle = flexTextRaised
      Next
      
      Flex.ColWidth(0) = 2151
      Flex.ColWidth(1) = 585
      Flex.ColWidth(2) = 675
      Flex.ColWidth(3) = 645
      
      Flex.ColAlignment(1) = flexAlignLeftCenter
      Flex.ColAlignment(2) = flexAlignLeftCenter
      Flex.ColAlignment(3) = flexAlignLeftCenter
      
   End With
   
   ColSortAsc(0) = True
   ColSortAsc(1) = False
   ColSortAsc(2) = True
   ColSortAsc(3) = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveSetting App.Title, "Settings", "txtServerIp", txtServerIP
   SaveSetting App.Title, "Settings", "txtServerPort", txtServerPort
   SaveSetting App.Title, "Settings", "chkAutoUpdate", chkAutoUpdate
   SaveSetting App.Title, "Settings", "chkNotify", chkNotify
End Sub

Private Sub Picture1_Click()
   frmAbout.Show vbModal
End Sub

Private Sub tmrUpdate_Timer()
   cmdQuery_Click
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
   
   On Error GoTo ErrHandler
   
   If PingTime = 0 Then PingTimer_End = Timer
   
   
   If (enumQueryState = dsNotQuerying) Then Exit Sub
   
   Dim sNewData As String
   Winsock1.GetData sNewData
   
   ReDim Preserve Packets(UBound(Packets) + 1)
   Packets(UBound(Packets)) = sNewData
   
   If InStr(1, sNewData, "\final\") Then
      bFinalPacketIsIn = True
      sFinalPacket = sNewData
      If AllPacketsAreHere(sNewData) Then enumQueryState = dsNotQuerying
   ElseIf bFinalPacketIsIn Then
      If AllPacketsAreHere(sFinalPacket) Then enumQueryState = dsNotQuerying
   End If
   
   Exit Sub
   
ErrHandler:
   
   Select Case Err
   Case 10054 'The connection is reset by remote side
      Resume
   Case Else
      MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & Err.Source, "ERROR in frmServerQuery.Winsock1_DataArrival()"
      enumQueryState = dsNotQuerying
   End Select
   
End Sub

Private Function AllPacketsAreHere(ByVal FinalPacket As String)
   
   Dim QueryNum As String, QueryNumPos As Long
   
   QueryNumPos = InStr(InStr(1, FinalPacket, "\queryid\"), FinalPacket, ".")
   QueryNum = Mid(FinalPacket, QueryNumPos + 1, InStr(QueryNumPos, FinalPacket, "\") - QueryNumPos - 1)
   If UBound(Packets) = QueryNum Then AllPacketsAreHere = True

End Function

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   MsgBox Number & vbCrLf & Description, vbCritical, "Winsock Error"
End Sub
