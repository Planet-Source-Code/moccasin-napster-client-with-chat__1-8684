VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00FF8080&
   Caption         =   "Napster Search"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9240
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TmrPing 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3720
      Top             =   2880
   End
   Begin VB.TextBox Pingresponse 
      Height          =   285
      Left            =   4560
      TabIndex        =   48
      Top             =   6360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4800
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4080
      Top             =   2880
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   4440
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Ping User"
      Height          =   375
      Left            =   2640
      TabIndex        =   46
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   42
      Top             =   5640
      Width           =   1095
   End
   Begin VB.ListBox LstFileName2 
      Height          =   255
      Left            =   3000
      TabIndex        =   41
      Top             =   6360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton CmdDownload 
      Caption         =   "&Download File"
      Height          =   375
      Left            =   0
      TabIndex        =   40
      Top             =   5640
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   1560
      TabIndex        =   26
      Top             =   6360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox TxtResults 
      Height          =   285
      Left            =   240
      TabIndex        =   25
      Top             =   6360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      Caption         =   "Search Results"
      Height          =   2895
      Left            =   0
      TabIndex        =   18
      Top             =   2760
      Width           =   9255
      Begin VB.ListBox LstFileName 
         Height          =   2205
         ItemData        =   "napstasearch.frx":0000
         Left            =   120
         List            =   "napstasearch.frx":0002
         TabIndex        =   47
         Top             =   480
         Width           =   2535
      End
      Begin VB.ListBox LstIPaddress 
         Height          =   2205
         Left            =   7800
         TabIndex        =   44
         Top             =   480
         Width           =   1335
      End
      Begin VB.ListBox LstLinkType 
         Height          =   2205
         Left            =   6960
         TabIndex        =   32
         Top             =   480
         Width           =   855
      End
      Begin VB.ListBox LstNick 
         Height          =   2205
         Left            =   5760
         TabIndex        =   31
         Top             =   480
         Width           =   1215
      End
      Begin VB.ListBox LstLength 
         Height          =   2205
         Left            =   5040
         TabIndex        =   30
         Top             =   480
         Width           =   735
      End
      Begin VB.ListBox LstFrequency 
         Height          =   2205
         Left            =   4080
         TabIndex        =   29
         Top             =   480
         Width           =   975
      End
      Begin VB.ListBox LstBitrate 
         Height          =   2205
         Left            =   3360
         TabIndex        =   28
         Top             =   480
         Width           =   735
      End
      Begin VB.ListBox LstSize 
         Height          =   2205
         ItemData        =   "napstasearch.frx":0004
         Left            =   2640
         List            =   "napstasearch.frx":0006
         TabIndex        =   27
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label17 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "IP address"
         Height          =   255
         Left            =   7800
         TabIndex        =   43
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label16 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Link type"
         Height          =   255
         Left            =   6960
         TabIndex        =   39
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label15 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nick"
         Height          =   255
         Left            =   5760
         TabIndex        =   38
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Length"
         Height          =   255
         Left            =   5040
         TabIndex        =   37
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Frequency"
         Height          =   255
         Left            =   4080
         TabIndex        =   36
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bitrate"
         Height          =   255
         Left            =   3360
         TabIndex        =   35
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Size"
         Height          =   255
         Left            =   2640
         TabIndex        =   34
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblfilename 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "File name"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Search info"
      Height          =   2535
      Left            =   0
      TabIndex        =   9
      Top             =   120
      Width           =   8055
      Begin VB.CommandButton Command2 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Find it!"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Caption         =   "Search for:"
         Height          =   1935
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   2415
         Begin VB.TextBox TxtArtist 
            Height          =   285
            Left            =   120
            TabIndex        =   1
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox TxtSong 
            Height          =   285
            Left            =   120
            TabIndex        =   2
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox TxtMax 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1080
            TabIndex        =   3
            Text            =   "100"
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Artist"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Song"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "Max Results:"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1560
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Advanced Search Options (Optional)"
         Height          =   1575
         Left            =   2640
         TabIndex        =   10
         Top             =   480
         Width           =   5295
         Begin VB.ComboBox CboValFrequency 
            Height          =   315
            ItemData        =   "napstasearch.frx":0008
            Left            =   3000
            List            =   "napstasearch.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1080
            Width           =   1695
         End
         Begin VB.ComboBox CboValBitrate 
            Height          =   315
            Left            =   3000
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   720
            Width           =   1695
         End
         Begin VB.ComboBox CboValLinespeed 
            Height          =   315
            Left            =   3000
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   360
            Width           =   1695
         End
         Begin VB.ComboBox CboBitrate 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   720
            Width           =   1815
         End
         Begin VB.ComboBox CboFrequency 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1080
            Width           =   1815
         End
         Begin VB.ComboBox CboLinespeed 
            Height          =   315
            ItemData        =   "napstasearch.frx":000C
            Left            =   1080
            List            =   "napstasearch.frx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label10 
            Caption         =   "HZ"
            Height          =   255
            Left            =   4800
            TabIndex        =   24
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label Label9 
            Caption         =   "Kb/s"
            Height          =   255
            Left            =   4800
            TabIndex        =   23
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label8 
            Caption         =   "K"
            Height          =   255
            Left            =   4800
            TabIndex        =   22
            Top             =   480
            Width           =   255
         End
         Begin VB.Label Label5 
            Caption         =   "Line Speed: "
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Bit rate:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "Frequency"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1080
            Width           =   855
         End
      End
   End
   Begin VB.Label Lbldownload 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   45
      Top             =   5760
      Width           =   5055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Server information"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6000
      Width           =   9255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bytelen As String, Napresults() As String
Dim Mp3path As String, Step2 As Integer, FileOpen As Boolean
Dim FileSize(1000) As String, Step3 As Integer, Pingtime As Long

Private Sub CmdDownload_Click()
Dim Nick As String, Filename As String
If LstFileName.ListIndex > -1 And LstFileName.ListCount > 0 Then
LstFileName2.ListIndex = LstFileName.ListIndex

'DOWNLOAD PATH OF MP3 ...
Mp3path = Environ("WinDir") & "\Desktop\" & Trim(LstFileName.List(LstFileName.ListIndex))

 If Trim(Mp3path) <> "" Then
    Nick = Trim(LstNick.List(LstNick.ListIndex))
    Filename = Trim(Chr(34) & LstFileName2.List(LstFileName2.ListIndex) & Chr(34))
    bytelen = Len(Nick & " " & Filename)
    
    If Form1!Winsock1.State = sckConnected And Nick <> "" And Filename <> "" Then
        Form1!Txtinfo = Form1!Txtinfo & vbCrLf & "attempting to download file..."
        Form1!Winsock1.SendData Chr(bytelen) & Chr(0)
        Form1!Winsock1.SendData Chr(203) & Chr(0)
        Form1!Winsock1.SendData Nick & " " & Filename
        CmdDownload.Enabled = False
        Frame4.Enabled = False
        Command3.Enabled = True
    Else
        MsgBox "Not connected to server!", vbExclamation, "Client error!"
    End If
    
  End If

End If
End Sub

Private Sub Command1_Click()
Dim Searchfor As String
Searchfor = ""
' the following code puts the search string in proper format
' before sending to server

If Not Val(TxtMax) > 0 Then
    MsgBox "Must have a max value", vbExclamation, "Client error!"
    Exit Sub
End If

If Trim(TxtArtist) = "" And Trim(TxtSong) = "" Then
MsgBox "Must have something to search!", vbExclamation, "Client error!"
Exit Sub
End If

If Trim(TxtArtist) <> "" Then
Searchfor = "FILENAME CONTAINS " & Chr(34) & TxtArtist & Chr(34)
End If

If Val(TxtMax) > 0 Then
Searchfor = Searchfor & " MAX_RESULTS " & CStr(Val(TxtMax))
End If

If Trim(TxtSong) <> "" Then
Searchfor = Searchfor & " FILENAME CONTAINS " & Chr(34) & TxtSong & Chr(34)
End If

If CboLinespeed.ListIndex <> -1 And CboValLinespeed.ListIndex <> -1 Then
    If CboLinespeed.List(CboLinespeed.ListIndex) <> "" And CboValLinespeed.List(CboValLinespeed.ListIndex) <> "" Then
    Searchfor = Searchfor & " LINESPEED " & Chr(34) & CboLinespeed.List(CboLinespeed.ListIndex) & Chr(34) & " " & CStr(CboValLinespeed.ListIndex)
    End If
End If

If CboBitrate.ListIndex <> -1 And CboValBitrate.ListIndex <> -1 Then
    If CboBitrate.List(CboBitrate.ListIndex) <> "" And CboValBitrate.List(CboValBitrate.ListIndex) <> "" Then
    Searchfor = Searchfor & " BITRATE " & Chr(34) & CboBitrate.List(CboBitrate.ListIndex) & Chr(34) & " " & Chr(34) & CboValBitrate.List(CboValBitrate.ListIndex) & Chr(34)
    End If
End If

If CboFrequency.ListIndex <> -1 And CboValFrequency.ListIndex <> -1 Then
    If CboFrequency.List(CboFrequency.ListIndex) <> "" And CboValFrequency.List(CboValFrequency.ListIndex) <> "" Then
    Searchfor = Searchfor & " FREQ " & Chr(34) & CboFrequency.List(CboFrequency.ListIndex) & Chr(34) & " " & Chr(34) & CboValFrequency.List(CboValFrequency.ListIndex) & Chr(34)
    End If
End If

TxtResults = ""
List1.Clear
LstFileName.Clear
LstFileName2.Clear
LstSize.Clear
LstBitrate.Clear
LstFrequency.Clear
LstLength.Clear
LstNick.Clear
LstLinkType.Clear
LstIPaddress.Clear

bytelen = Len(Searchfor)

If Form1!Winsock1.State = sckConnected Then
Form1!Txtinfo = Form1!Txtinfo & vbCrLf & "Sending search query..."
Form1!Winsock1.SendData Chr(bytelen) & Chr(0)
Form1!Winsock1.SendData Chr(200) & Chr(0)
Form1!Winsock1.SendData Searchfor
Else
MsgBox "Not connected to server!", vbExclamation, "Client Error"
End If

End Sub

Private Sub Command2_Click()
TxtArtist = ""
TxtSong = ""
End Sub

Private Sub Command3_Click()
'If Form1!Winsock1.State = sckConnected Then
'Form1!Winsock1.SendData Chr(0)
'Form1!Winsock1.SendData Chr(0)
'Form1!Winsock1.SendData Chr(219) & Chr(0)
'End If
Timer1.Enabled = False

Winsock2.Close
While Winsock2.State
  DoEvents
Wend

Winsock1.Close
While Winsock1.State
DoEvents
Wend

CmdDownload.Enabled = True
Command3.Enabled = False
Frame4.Enabled = True
Lbldownload.Caption = "Download Cancelled"
End Sub

Private Sub Command4_Click()
Dim bytelen As Long
If LstNick.ListIndex > -1 Then
If Form1!Winsock1.State = sckConnected Then
bytelen = Len(LstNick.List(LstNick.ListIndex))
Form1!Winsock1.SendData Chr(bytelen) & Chr(0)
Form1!Winsock1.SendData Chr(239) & Chr(2)
Form1!Winsock1.SendData Trim(LstNick.List(LstNick.ListIndex))
Pingtime = 0
TmrPing.Enabled = True
End If
End If
End Sub

Private Sub Form_Load()
Step2 = 0
Step3 = 0
FileOpen = False

CboLinespeed.AddItem ""
CboLinespeed.AddItem "AT LEAST"
CboLinespeed.AddItem "AT BEST"
CboLinespeed.AddItem "EQUAL TO"

CboValLinespeed.AddItem ""
CboValLinespeed.AddItem "14.4 kbps"
CboValLinespeed.AddItem "28.8 kbps"
CboValLinespeed.AddItem "33.6 kbps"
CboValLinespeed.AddItem "56.7 kbps"
CboValLinespeed.AddItem "64K ISDN"
CboValLinespeed.AddItem "128K ISDN"
CboValLinespeed.AddItem "Cable"
CboValLinespeed.AddItem "DSL"
CboValLinespeed.AddItem "T1"
CboValLinespeed.AddItem "T3 +"

CboBitrate.AddItem ""
CboBitrate.AddItem "AT LEAST"
CboBitrate.AddItem "AT BEST"
CboBitrate.AddItem "EQUAL TO"

CboValBitrate.AddItem ""
CboValBitrate.AddItem "256"
CboValBitrate.AddItem "192"
CboValBitrate.AddItem "160"
CboValBitrate.AddItem "128"
CboValBitrate.AddItem "112"
CboValBitrate.AddItem "98"
CboValBitrate.AddItem "64"
CboValBitrate.AddItem "56"
CboValBitrate.AddItem "48"
CboValBitrate.AddItem "32"
CboValBitrate.AddItem "24"
CboValBitrate.AddItem "20"


CboFrequency.AddItem ""
CboFrequency.AddItem "AT LEAST"
CboFrequency.AddItem "AT BEST"
CboFrequency.AddItem "EQUAL TO"

CboValFrequency.AddItem ""
CboValFrequency.AddItem "48000"
CboValFrequency.AddItem "44100"
CboValFrequency.AddItem "32000"
CboValFrequency.AddItem "24000"
CboValFrequency.AddItem "22050"
CboValFrequency.AddItem "16000"
CboValFrequency.AddItem "12000"
CboValFrequency.AddItem "11025"
CboValFrequency.AddItem "8000"

End Sub



Private Sub Form_Unload(Cancel As Integer)
Close
End Sub

Private Sub LstBitrate_Click()
If LstBitrate.ListIndex > -1 And LstBitrate.ListCount > 0 Then
LstSize.ListIndex = LstBitrate.ListIndex
LstFileName.ListIndex = LstBitrate.ListIndex
LstFrequency.ListIndex = LstBitrate.ListIndex
LstLength.ListIndex = LstBitrate.ListIndex
LstNick.ListIndex = LstBitrate.ListIndex
LstLinkType.ListIndex = LstBitrate.ListIndex
LstIPaddress.ListIndex = LstBitrate.ListIndex
End If
End Sub

Private Sub LstFileName_Click()
If LstFileName.ListIndex > -1 And LstFileName.ListCount > 0 Then
LstSize.ListIndex = LstFileName.ListIndex
LstBitrate.ListIndex = LstFileName.ListIndex
LstFrequency.ListIndex = LstFileName.ListIndex
LstLength.ListIndex = LstFileName.ListIndex
LstNick.ListIndex = LstFileName.ListIndex
LstLinkType.ListIndex = LstFileName.ListIndex
LstIPaddress.ListIndex = LstFileName.ListIndex
End If
End Sub

Private Sub LstFrequency_Click()
If LstFrequency.ListIndex > -1 And LstFrequency.ListCount > 0 Then
LstSize.ListIndex = LstFrequency.ListIndex
LstBitrate.ListIndex = LstFrequency.ListIndex
LstFileName.ListIndex = LstFrequency.ListIndex
LstLength.ListIndex = LstFrequency.ListIndex
LstNick.ListIndex = LstFrequency.ListIndex
LstLinkType.ListIndex = LstFrequency.ListIndex
LstIPaddress.ListIndex = LstFrequency.ListIndex
End If
End Sub

Private Sub LstIPaddress_Click()
If LstIPaddress.ListIndex > -1 And LstIPaddress.ListCount > 0 Then
LstSize.ListIndex = LstIPaddress.ListIndex
LstBitrate.ListIndex = LstIPaddress.ListIndex
LstFrequency.ListIndex = LstIPaddress.ListIndex
LstLength.ListIndex = LstIPaddress.ListIndex
LstNick.ListIndex = LstIPaddress.ListIndex
LstLinkType.ListIndex = LstIPaddress.ListIndex
LstFileName.ListIndex = LstIPaddress.ListIndex
End If
End Sub

Private Sub LstLength_Click()
If LstLength.ListIndex > -1 And LstSize.ListCount > 0 Then
LstSize.ListIndex = LstLength.ListIndex
LstBitrate.ListIndex = LstLength.ListIndex
LstFrequency.ListIndex = LstLength.ListIndex
LstFileName.ListIndex = LstLength.ListIndex
LstNick.ListIndex = LstLength.ListIndex
LstLinkType.ListIndex = LstLength.ListIndex
LstIPaddress.ListIndex = LstLength.ListIndex
End If
End Sub

Private Sub LstLinkType_Click()
If LstLinkType.ListIndex > -1 And LstLinkType.ListCount > 0 Then
LstSize.ListIndex = LstLinkType.ListIndex
LstBitrate.ListIndex = LstLinkType.ListIndex
LstFrequency.ListIndex = LstLinkType.ListIndex
LstLength.ListIndex = LstLinkType.ListIndex
LstNick.ListIndex = LstLinkType.ListIndex
LstFileName.ListIndex = LstLinkType.ListIndex
LstIPaddress.ListIndex = LstLinkType.ListIndex
End If
End Sub

Private Sub LstNick_Click()
If LstNick.ListIndex > -1 And LstNick.ListCount > 0 Then
LstSize.ListIndex = LstNick.ListIndex
LstBitrate.ListIndex = LstNick.ListIndex
LstFrequency.ListIndex = LstNick.ListIndex
LstLength.ListIndex = LstNick.ListIndex
LstFileName.ListIndex = LstNick.ListIndex
LstLinkType.ListIndex = LstNick.ListIndex
LstIPaddress.ListIndex = LstNick.ListIndex
End If
End Sub

Private Sub LstSize_Click()
If LstSize.ListIndex > -1 And LstSize.ListCount > 0 Then
LstFileName.ListIndex = LstSize.ListIndex
LstBitrate.ListIndex = LstSize.ListIndex
LstFrequency.ListIndex = LstSize.ListIndex
LstLength.ListIndex = LstSize.ListIndex
LstNick.ListIndex = LstSize.ListIndex
LstLinkType.ListIndex = LstSize.ListIndex
LstIPaddress.ListIndex = LstSize.ListIndex
End If
End Sub



Private Sub Pingresponse_Change()
MsgBox "Ping response for " & Pingresponse.Text & " is " & CStr(Pingtime) & " ms"
TmrPing.Enabled = False
End Sub

Private Sub Timer1_Timer()
    Lbldownload.Caption = CStr(LOF(1) * 100 \ CLng(Trim(FileSize(LstSize.ListIndex)))) & "% downloaded"
    DoEvents
End Sub

Private Sub TmrPing_Timer()
If Pingtime < 500 Then
Pingtime = Pingtime + 1
Else
MsgBox "Ping request timed out!", vbInformation, "Client ping"
End If
End Sub

Private Sub TxtResults_Change()
Dim stp
If InStr(1, TxtResults, Chr(202), vbTextCompare) <> 0 Then
TxtResults = Replace(TxtResults, Chr(202), "")
Napresults() = Split(TxtResults, Chr(201))
For stp = 0 To UBound(Napresults)
List1.AddItem Napresults(stp)
Next stp
Listnapresults
End If
End Sub

Sub Listnapresults()
Dim stp
Dim A As String, B As String, C() As String
On Error Resume Next
' formats results
If List1.ListCount > 0 Then
    For stp = 0 To (List1.ListCount - 1)
        If Trim(List1.List(stp)) <> "" Then
        A = List1.List(stp)
        B = inbetween(A, Chr(34), Chr(34))
        A = GetLast(B, "\")
        If Len(Trim(A)) <> 0 And Len(Trim(B)) <> 0 Then
        LstFileName.AddItem A
        LstFileName2.AddItem B
        End If
        B = Trim(GetLast(List1.List(stp), Chr(34)))
        B = Trim(afterit(" ", B))
        C() = Split(B, " ")
        If UBound(C) = 6 Then
        LstSize.AddItem Format(CStr((Val(C(0)) / 1048576)), "##.00") & " MB"
        FileSize(LstSize.ListCount - 1) = C(0)
        LstBitrate.AddItem (C(1)) & " Kb/s"
        LstFrequency.AddItem C(2) & " HZ"
        B = Replace(Format(CStr(Val(C(3)) / 60), "#.00"), ".", ":")
        A = Format(beforeit(":", B), "0") & ":" & Format(CStr(Val(afterit(":", B) * 6 / 10)), "00")
        LstLength.AddItem A
        LstNick.AddItem C(4)
        LstIPaddress.AddItem IPToString(Val(C(5)))
        If Val(C(6)) = 0 Then
        C(6) = "Unknown"
        ElseIf Val(C(6)) = 1 Then
        C(6) = "14.4 K"
        ElseIf Val(C(6)) = 2 Then
        C(6) = "28.8 K"
        ElseIf Val(C(6)) = 3 Then
        C(6) = "33.6 K"
        ElseIf Val(C(6)) = 4 Then
        C(6) = "56.7 K"
        ElseIf Val(C(6)) = 5 Then
        C(6) = "64 K ISDN"
        ElseIf Val(C(6)) = 6 Then
        C(6) = "126 K ISDN"
        ElseIf Val(C(6)) = 7 Then
        C(6) = "Cable"
        ElseIf Val(C(6)) = 8 Then
        C(6) = "DSL"
        ElseIf Val(C(6)) = 9 Then
        C(6) = "T1"
        ElseIf Val(C(6)) = 10 Then
        C(6) = "T3 +"
        End If
        LstLinkType.AddItem C(6)
        End If
        End If
    Next stp
Else
LstFileName.AddItem "No matches found!"
End If
End Sub


Private Sub Winsock1_Close()
If Form1!Winsock1.State = sckConnected Then
Form1!Winsock1.SendData Chr(0)
Form1!Winsock1.SendData Chr(0)
Form1!Winsock1.SendData Chr(219) & Chr(0) 'sends download complete
End If

Timer1.Enabled = False
CmdDownload.Enabled = True
Command3.Enabled = False
Frame4.Enabled = True
FileOpen = False
Step3 = 0
Close
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
If (Winsock1.State <> sckClosed) Then Winsock1.Close ' If not closed then call close method to cleanup socket status
Winsock1.Accept requestID ' Accept the incoming connection
Winsock1.SendData Chr(49) 'send "1"
Step = 0
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim SockData As String, Parsed() As String
Winsock1.GetData SockData
Select Case Step3
Case 0
SockData = Replace(SockData, Chr(0), "")
    If InStr(1, SockData, "SEND", vbTextCompare) <> 0 Then
        If afterit("SEND", SockData) = "" Then
        Step3 = Step3 + 1
        ElseIf InStr(1, LCase(Trim(SockData)), LCase(Trim(Form1!Text1.Text)), vbTextCompare) <> 0 Then
         If LCase(inbetween(SockData, Chr(34), Chr(34))) Like LCase(LstFileName2.List(LstFileName2.ListIndex)) Then
         Winsock1.SendData "0"
         Step3 = Step3 + 2
         End If
        Else
        Winsock1.SendData "INVALID REQUEST"
        End If
    End If
Case 1
    SockData = Replace(SockData, Chr(0), "")
    If InStr(1, LCase(Trim(SockData)), LCase(Trim(Form1!Text1.Text)), vbTextCompare) Then
     If LCase(inbetween(SockData, Chr(34), Chr(34))) Like LCase(LstFileName2.List(LstFileName2.ListIndex)) Then
         Winsock1.SendData "0"
         Step3 = Step3 + 1
     End If
    Else
    Winsock1.SendData "INVALID REQUEST"
    End If
Case 2
        SockData = Replace(SockData, Chr(0), "")
    If InStr(1, SockData, Chr(255) & Chr(255) & Chr(255), vbTextCompare) Then
        If FileOpen = False Then
        Open Mp3path For Binary Access Write As #1
        FileOpen = True
        Put 1, , SockData
        Winsock2.SendData Chr(0)
        Winsock2.SendData Chr(0)
        Winsock2.SendData Chr(218) & Chr(0)
        Lbldownload = "Download Started"
        'lets server know we are downloading
        Timer1.Enabled = True
        Step3 = Step3 + 1
        Else
        MsgBox SockData
        End If
    End If
Case 3
    ' (thanks, jim, for the lof suggestion)
    If CLng(LOF(1)) >= CLng(FileSize(LstSize.ListIndex)) Then
    'file d/l complete
    FileOpen = False
    Close #1
    Form1!Winsock1.SendData Chr(0)
    Form1!Winsock1.SendData Chr(0)
    Form1!Winsock1.SendData Chr(219) & Chr(0) 'tells server d/l is complete
    Timer1.Enabled = False
    MsgBox "Download complete!", vbInformation, "Download info."
    Else
        If FileOpen = True Then
            Put 1, , SockData
        End If
    End If
    
End Select
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox Description, vbExclamation, "Client error!"
End Sub

Private Sub Winsock2_Close()
Form1!Txtinfo = Form1!Txtinfo & vbCrLf & "Download cancelled"

If Form1!Winsock1.State = sckConnected Then
Form1!Winsock1.SendData Chr(0)
Form1!Winsock1.SendData Chr(0)
Form1!Winsock1.SendData Chr(219) & Chr(0) 'sends download complete
End If

Close

Timer1.Enabled = False
CmdDownload.Enabled = True
Command3.Enabled = False
Frame4.Enabled = True
FileOpen = False
Step2 = 0
End Sub


Private Sub Winsock2_Connect()
Lbldownload.Caption = "Connected."
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
Dim SockData As String
Winsock2.GetData SockData
Select Case Step2
Case 0
    SockData = Replace(SockData, Chr(0), "")
    If InStr(1, SockData, Chr(49), vbTextCompare) <> 0 Then
        Winsock2.SendData "GET"
        Pause 0.51
        Winsock2.SendData Trim(Form1!Text1) & " " & Chr(34) & LstFileName2.List(LstFileName2.ListIndex) & Chr(34) & " 0"
        Step2 = Step2 + 1
    End If
Case 1
    SockData = Replace(SockData, Chr(0), "")
    If InStr(1, SockData, Chr(255) & Chr(255) & Chr(255), vbTextCompare) Then
        If FileOpen = False Then
        Open Mp3path For Binary Access Write As #1
        FileOpen = True
        Put 1, , SockData
        Winsock2.SendData Chr(0)
        Winsock2.SendData Chr(0)
        Winsock2.SendData Chr(218) & Chr(0)
        Lbldownload.Caption = "Download Started."
        Form1!Txtinfo = Form1!Txtinfo & vbCrLf & "Download started."
        Form1!Txtinfo = Form1!Txtinfo & vbCrLf & "Downloading to " & Mp3path
        'lets server know we are downloading
        Timer1.Enabled = True
        Step2 = Step2 + 1
        Else
        MsgBox SockData
        End If
    End If
Case 2
    If LOF(1) >= CLng(FileSize(LstSize.ListIndex)) Then
    Close #1
    Form1!Winsock1.SendData Chr(0)
    Form1!Winsock1.SendData Chr(0)
    Form1!Winsock1.SendData Chr(219) & Chr(0) 'sends download complete
    Timer1.Enabled = False
    MsgBox "Download complete!", vbInformation, "Client info"
    Else
        If FileOpen = True Then
            Put 1, , SockData
        End If
    End If
End Select
End Sub

Private Sub Winsock2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox Description, vbExclamation, "Client error!"
Form1!Txtinfo = Form1!Txtinfo & vbCrLf & "Error: " & Description
End Sub

