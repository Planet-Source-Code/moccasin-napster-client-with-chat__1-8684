VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   Caption         =   "Napster Client by Moccasin"
   ClientHeight    =   6270
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3600
      Top             =   2880
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Disconnect"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2520
      Top             =   2760
   End
   Begin VB.Frame Frame3 
      Caption         =   "Status Area"
      Height          =   2295
      Left            =   2760
      TabIndex        =   7
      Top             =   120
      Width           =   5415
      Begin VB.TextBox Txtinfo 
         Height          =   1935
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Napster MOTD"
      Height          =   3375
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   8055
      Begin VB.TextBox Text4 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   3015
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   240
         Width           =   7815
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2040
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Connect"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Napster info."
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "napsta.frx":0000
         Left            =   1080
         List            =   "napsta.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   480
         TabIndex        =   3
         Text            =   "6699"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "Username"
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Port"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   375
      End
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Server information."
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   6000
      Width           =   8295
   End
   Begin VB.Menu mnuoption 
      Caption         =   "&Options"
      Begin VB.Menu mnusearch 
         Caption         =   "&Search for song"
      End
      Begin VB.Menu mnuchat 
         Caption         =   "&Chat"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'     *****************************************
'     *** Napster Client coded by Moccasin ****
'     *****************************************
'               Update June 6, 2000

'         Started May 27 | Released June 1

' Keep in mind this is *merely an example*
' for educational purposes.


' Yes, this client is  functional (it can search, d/l, and ping)
' but it is still limited compared to the official
' windows client (this doesn't support multiple d/l's,
' hotlists, resume downloads, etc).

'****** Known Issues **********************

' 1. there is a problem with the firewall downloading.
'    I have not clearly determined whether or not this is a
'    bug or a server/remote client issue, or perhaps both.
'    When firewall downloading, sometimes napster disconnects
'    and you end up getting the "Connection reset by remote.." error
'    luckily many users are be running
'    a napster listening port  w/o a firewall.

' 2. when a download is at 99% downloaded, the remote client
'    will sometimes disconnect you (you will then
'    get a "Connection reset by remote side." msgbox). But the
'    lost bytes will only result in a second or two shortening
'    of the mp3 length.

'    I'm not sure if this disconnection is supposed to happen
'    or not, because the protocol spec didn't specify exactly
'    who is supposed to end the transfer between clients
'    and how.

' 3. if an incoming packet contains more than one chat
'    data/message (which is somewhat rare),
'    then a few weird characters may appear on
'    the chat screen. This is only a tiny bug

' 4. this program uses the vb6 functions "split" (for parsing strings)
'    and "replace" (for replacing strings). If you are using vb5,
'    you may find the equivalent of these functions at PSC
'    www.planet-source-code.com/vb

' ************ Outline of this code ********

' I know the code is a bit confusing
' Anyway, this is how it should work...

' 1. Connecting. You put your napster username in text1, password
'    in text2. Pressing "connect.." causes winsock1 to
'    connect to the main napster server (server.napster.com),
'    as soon as it's connected, napster sends an
'    available napster server address with port.
'    Winsock1 connects to the available server and
'    subsequently sends the login information. The server
'    should respond with a 0x03 if the login was a success.

' 2. Searching for an mp3. Fill in search fields. press
'    "Find it". This will cause winsock1 on form1 to
'    search for the query. The results are then stored in
'    txtresults on form2. When the Txtresults textbox see's
'    there's an "end search" character from the server it
'    parses the search and loads each field into the appropriate
'    list boxes.


'  3. Downloading. Choose a mp3 to download from the list box.
'     Press "download", this will cause winsock1 on form1
'     to request a file from a client. The server should
'     then return the IP address of the client along
'     with the port to connect to. If the port is not 0, then
'     this means that it is not firewalled. However if it is,
'     then it is firewalled and a different method of download
'     is required. If it is not firewalled, winsock2 on form2
'     will request the file from the remote client. When the
'     file is sent it will be stored under the same name in
'     the windows desktop.

'     If the port is firewalled (i.e. "0") then this either
'     means 1) it actually is firewalled 2) the client is
'     misconfigured 3) the user is stingy and greedy.
'     The code I originally used for firewalled downloading
'     seemed to be buggy, so instead, when the port number
'     returns "0", it just pops up a msgbox stating so..

'     Though, I left the firewall downloading code there for example
'     purposes of how it would be done.

' *****************************************
' This project was tested on win 98 w/ a 56k modem. Works
' fine here.


' ***** Resources **************************
' Email: to_moccasin@hotmail.com
' OpenNap Project: http://opennap.sourceforge.net/
' Napster protocol spec: http://opennap.sourceforge.net/napster.txt
' Napster dev. mailing list site: http://www.egroups.com/community/napdev

' To better understand this code, please check out the
' napster protocol spec at the aforementioned URL


' ***** Other Stuff ************************
' How to send the data type field within vb.

' For data type number N, in which N falls inbetween dec numbers
' 1 to 255. You may send N as follows: chr(N) & chr(0)

' However, if N is greater than 255, such as data type number
' 400 (which is the join a chat request) you would need to
' find the dec equivalent to this number. What I like to do is
' convert it to hex, then from there look it up in my conversion
' chart to find the dec conversion (there are easier ways of
' doing this though).

' For example, 400 in hex is (0x90, 0x01). This can be found
' simply using the hex function in vb (i.e. msgbox hex(400)
' will say "190" - 0x90, 0x01).

' What I would then do is look up the 0x90 part since hex
' 0x01 is the same number in dec. You would find that hex
' 90 is dec 144. Thus, I would send this data field as
' "chr(144) & chr(1)"

' You can find a hex to dec conversion chart at
' http://logjam.nerdc.ufl.edu/hex.html


' Email me if you still have any questions.



Dim Napinfo
Dim Step As Integer, Secnd As Integer, Secnd2 As Integer
Dim bytelen As Long, StrNameList As String
Dim Naplogon As Boolean, Napresponse As String, StrChatlist As String

Private Sub Command1_Click()
    MainConnect ' connect to main napster server
End Sub

Private Sub Command2_Click()
    NapDisconnect ' disconnect
End Sub




Private Sub Form_Load()
' Combo1 is the list for your connection speed
    Combo1.AddItem "unknown"
    Combo1.AddItem "14.4 Kb/s"
    Combo1.AddItem "28.8 Kb/s"
    Combo1.AddItem "33.6 Kb/s"
    Combo1.AddItem "56.7 Kb/s"
    Combo1.AddItem "64K ISDN"
    Combo1.AddItem "128K ISDN"
    Combo1.AddItem "Cable"
    Combo1.AddItem "DSL"
    Combo1.AddItem "T1"
    Combo1.AddItem "T3 +"
    Combo1.Text = Combo1.List(4)

' sets "false" for whether or not you are logged onto napster
' yet
    Naplogon = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
' close all winsock controls before unloading
Winsock1.Close
    While Winsock1.State
        DoEvents
    Wend
    
Form2!Winsock2.Close
    While Form2!Winsock2.State
        DoEvents
    Wend
    
Form2!Winsock1.Close
    While Form2!Winsock1.State
        DoEvents
    Wend
End
End Sub

Private Sub Label2_Change()
    Form2!Label3 = Label2
End Sub

Private Sub mnuchat_Click()
    If Winsock1.State = sckConnected Then
        '(617) request channel list - no data
        
        'if channel list recieved, then frmchatlist will
        ' pop up showing the list
        Winsock1.SendData Chr(0)
        Winsock1.SendData Chr(0)
        Winsock1.SendData Chr(105) & Chr(2) 'sends 617
        Secnd2 = 0
        Timer2.Enabled = True
    End If
End Sub

Private Sub mnusearch_Click()
    'shows the mp3 search form
    Form2.Show
End Sub

Private Sub Text4_Change()
On Error Resume Next
    Text4.SelLength = 0
    If Len(Text4.Text) > 0 Then
        If Right$(Text4.Text, 1) = vbCrLf Then
            Text4.SelStart = Len(Text4.Text) - 1
            Exit Sub
        End If
        Text4.SelStart = Len(Text4.Text)
    End If
End Sub

Private Sub Timer1_Timer()
Dim userresp
' sees if you are connected to the nap server yet,
' if you are not connected within a 20 seconds
' it will prompt to try for reconnect

    If Secnd >= 20 Then
        'should get connected to a napster server within 20 seconds
        userresp = MsgBox("Connection timed out!", vbRetryCancel, "Napster Error!")
        If userresp = vbRetry Then
                NapDisconnect
                MainConnect
        Else
                Txtinfo = Txtinfo & vbCrLf & "Connection timed out."
                NapDisconnect
        End If
        Secnd = 0
    End If

    If Naplogon = False Then
        Secnd = Secnd + 1
    Else
        Timer1.Enabled = False
    End If

End Sub



Private Sub Timer2_Timer()
Dim userresp 'user response

' sees if you have gotten the channel list
' if you have not recieved the list within a 10 seconds
' it will prompt to for retry

    If Secnd2 >= 10 Then
        'should get recieve chat list within 10 seconds
        userresp = MsgBox("Chat list request timed out!", vbRetryCancel, "Napster Error!")
        If userresp = vbCancel Then
                Timer2.Enabled = False
        End If
        Secnd2 = 0
    End If

End Sub

Private Sub Txtinfo_Change()
' this makes sure you see the newest text
On Error Resume Next
    Txtinfo.SelLength = 0
    If Len(Txtinfo.Text) > 0 Then
        If Right$(Txtinfo.Text, 1) = vbCrLf Then
            Txtinfo.SelStart = Len(Txtinfo.Text) - 1
            Exit Sub
        End If
        Txtinfo.SelStart = Len(Txtinfo.Text)
    End If
End Sub

Private Sub Winsock1_Close()
    Command1.Enabled = True
    Command2.Enabled = False
    Text1.Enabled = True
    Form2!CmdDownload.Enabled = True
    Form2!Command3.Enabled = False
    Form2!Frame4.Enabled = True
End Sub

Private Sub Winsock1_Connect()
' sends login info
    NapConnect
End Sub




Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim Altserver As String, AltIp As String, Parsed() As String, stp2
Dim AltPort As String, a As String, B As String, C As String, d As String
Dim napparsed() As String, napparsed2() As String, strlen As Long, napresp As String, stp
Dim napparsed3() As String
Select Case Step
Case 1
    Winsock1.GetData Napresponse
    ' the napresponse here will, or should be, the IP of the
    ' available nap server to connect to and the port as
    ' well
    
    ' set step to 2
    Step = Step + 1
    Altserver = Trim(Napresponse)
    Altserver = Replace(Altserver, Chr(10), "")
    Altserver = Replace(Altserver, Chr(13), "")
    AltIp = beforeit(":", Altserver) ' ip
    AltPort = afterit(":", Altserver) ' port
    Winsock1.Close
    While Winsock1.State
      DoEvents
    Wend

    Txtinfo = Txtinfo & vbCrLf & "Main Napster Server recommended: " & AltIp & " on port " & AltPort
    Txtinfo = Txtinfo & vbCrLf & "Connecting to: " & AltIp & " on port " & AltPort
    Winsock1.Connect AltIp, AltPort
    Napresponse = ""
Case 2
    
    Winsock1.GetData Napresponse
    
    'Napresponse = Replace(Napresponse, Chr(0), "") 'vb6 function
    'Napresponse = Replace(Napresponse, Chr(109) & Chr(2), vbCrLf)
    'if your email address is in the message with the "logon success" character
    If InStr(1, Napresponse, "@", vbTextCompare) <> 0 And InStr(1, Napresponse, Chr(3), vbTextCompare) <> 0 Then
        ' you user name and password was excepted
        ' you are now logged on
        Naplogon = True
        Txtinfo = Txtinfo & vbCrLf & "Successfully logged onto Napster!"
    Else
        If InStr(1, Napresponse, Chr(214), vbTextCompare) <> 0 And InStr(1, Napresponse, Chr(66), vbTextCompare) <> 0 Then
        ' you may also use the function parse for the string...

        a = Trim(inbetween(Napresponse, Chr(214), Chr(66)))
        B = beforeit(" ", a) 'users
        C = afterit(" ", a)
        d = beforeit(" ", C) 'number of mp3s available on server
        C = afterit(" ", C) ' GB of mp3's
        Label2 = "There are currently " & Format(d, "###,###") & " mp3's, totaling " & Format(C, "###,###") & " GB, available in " & Format(B, "###,###") & " libraries"
             
        Napresponse = afterit(Chr(66), Napresponse)
        
        If InStr(1, Napresponse, Chr(109) & Chr(2), vbTextCompare) <> 0 Then
        ' this will take out any extraneous characters..
        Napresponse = Replace(Napresponse, Chr(0), "")
        napparsed() = Split(Napresponse, Chr(109) & Chr(2))
            For stp = LBound(napparsed) To UBound(napparsed)
            strlen = Len(napparsed(stp)) - 1
            If strlen > 0 Then
            napresp = napresp & Left(napparsed(stp), strlen) & vbCrLf
            Else
            napresp = napresp & vbCrLf
            End If
            Next
        End If
        Text4 = napresp
        Else
        If InStr(1, Napresponse, Chr(109) & Chr(2), vbTextCompare) <> 0 Then
        Napresponse = Replace(Napresponse, Chr(0), "")
        napparsed() = Split(Napresponse, Chr(109) & Chr(2))
            For stp = LBound(napparsed) To UBound(napparsed)
            strlen = Len(napparsed(stp)) - 1
            If strlen > 0 Then
            napresp = napresp & Left(napparsed(stp), strlen) & vbCrLf
            Else
            napresp = napresp & vbCrLf
            End If
            Next
        End If
        Text4 = Text4 & napresp
        End If
        
        If InStr(1, Napresponse, "Want to promote", vbTextCompare) <> 0 Then
            'end of Message of the Day (MOTD)..proceed to next step
            
            Step = Step + 1
        End If
    Napresponse = ""
    End If
Case 3
Winsock1.GetData Napresponse
Napresponse = Replace(Napresponse, Chr(0), "")
    If InStr(1, Napresponse, Chr(214), vbTextCompare) <> 0 And InStr(1, Napresponse, Chr(147) & Chr(1), vbTextCompare) = 0 Then
        Napresponse = Trim(afterit(Chr(214), Napresponse))
        a = beforeit(" ", Napresponse) 'users (libraries)
        B = afterit(" ", Napresponse)
        C = beforeit(" ", B) 'number of files (mp3's)
        B = afterit(" ", B) 'size in gigabytes of particular server
        Label2 = "There are currently " & Format(C, "###,###") & " mp3's, totaling " & Format(B, "###,###") & " GB, available in " & Format(a, "###,###") & " libraries"
    
    ElseIf InStr(1, Napresponse, Chr(147) & Chr(1), vbTextCompare) Then
    'gets what a user  said
        On Error Resume Next
        If Trim(beforeit(" ", afterit(Chr(1), Napresponse))) = Frmchat.Caption Then
            Napresponse = afterit(Frmchat.Caption & " ", Napresponse)
            Napresponse = "<" & beforeit(" ", Napresponse) & "> " & afterit(" ", Napresponse)
            Frmchat!TxtChat = Frmchat!TxtChat & vbCrLf & Napresponse
        Else
            Frmchat.Caption = Trim(beforeit(" ", afterit(Chr(1), Napresponse)))
            Napresponse = afterit(" ", Napresponse)
            Napresponse = "<" & beforeit(" ", Napresponse) & "> " & afterit(" ", Napresponse)
            Frmchat!TxtChat = Frmchat!TxtChat & vbCrLf & Napresponse
        End If


    ElseIf InStr(1, Napresponse, Chr(204), vbTextCompare) <> 0 Then
        Parsed = Split(Napresponse, " ")
        a = Trim(Parsed(2)) ' port
        B = IPToString(Val(Trim(Parsed(1)))) ' ip
        If B <> "0.0.0.0" Then
            napparsed3() = Split(B, ".")
            If napparsed3(3) = "0" Then Exit Sub
            
            If Val(a) > 0 Then
                Form2!Lbldownload.Caption = "Connecting to remote client: " & B & " on port " & a
                Form2!Winsock2.Close
                While Form2!Winsock2.State
                    DoEvents
                Wend
                Form2!Winsock2.Connect B, a
            Else
                Txtinfo = Txtinfo & vbCrLf & "Probable firewalled port on remote client.."
                ' port is thus 0
                MsgBox "Can't download, probable firewalled port on remote client - port returned was " & a, vbExclamation, "Download Error"
                ' this original firewall d/l code is a bit buggy....
        
                'Form2!Winsock1.Close
                'While Form2!Winsock1.State
                'DoEvents
                'Wend
        
                'Form2!Winsock1.LocalPort = Trim(Text3.Text)
                'Form2!Winsock1.Listen
                'C = Trim(Form2!LstNick.List(Form2!LstNick.ListIndex)) & " " & Chr(34) & Trim(Form2!LstFileName2.List(Form2!LstFileName2.ListIndex)) & Chr(34)
                'bytelen = Len(C)
                'Winsock1.SendData bytelen & Chr(0)
                'Winsock1.SendData Chr(244) & Chr(1)
                'Winsock1.SendData C
            End If
        End If
    ElseIf InStr(1, Napresponse, Chr(106) & Chr(2), vbTextCompare) <> 0 Then
    ' returns channel list
    Timer2.Enabled = False
    Secnd2 = 0
    On Error Resume Next
        StrChatlist = StrChatlist & Napresponse
        If InStr(1, Napresponse, Chr(105) & Chr(2), vbTextCompare) <> 0 Then
            'Napresponse = Replace(Napresponse, Chr(0), "")
            napparsed() = Split(StrChatlist, Chr(106) & Chr(2))
            For stp = LBound(napparsed) To UBound(napparsed)
                strlen = Len(napparsed(stp)) - 1
                If strlen > 0 Then
                    napparsed2() = Split(Left(napparsed(stp), strlen), " ")
                    ' channel name  & number of people in it
                    If Len(napparsed2(1)) > 0 Then
                        If Len(napparsed2(0)) > 10 Then
                            Frmchatlist!List1.AddItem napparsed2(0) & Chr(9) & napparsed2(1)
                        Else
                            Frmchatlist!List1.AddItem napparsed2(0) & Chr(9) & Chr(9) & napparsed2(1)
                        End If
                    Else
                        Frmchatlist!List1.AddItem "Error! Try again"
                    End If
                End If
            Next
            'empty strchatlist and show frmchatlist if
            'chatlist was recieved
            StrChatlist = ""
            Frmchatlist.Show
        End If

    ElseIf InStr(1, Napresponse, Chr(105) & Chr(2), vbTextCompare) <> 0 Then
    On Error Resume Next
    'Napresponse = Replace(Napresponse, Chr(0), "")
    napparsed() = Split(StrChatlist, Chr(106) & Chr(2))
            For stp = LBound(napparsed) To UBound(napparsed)
                strlen = Len(napparsed(stp)) - 1
                If strlen > 0 Then
                    napparsed2() = Split(Left(napparsed(stp), strlen), " ")
                    Frmchatlist!List1.AddItem napparsed2(0)
                    Frmchatlist!List2.AddItem napparsed2(1)
                End If
            Next
            StrChatlist = ""
            Frmchatlist.Show

    ElseIf InStr(1, Napresponse, Chr(152) & Chr(1), vbTextCompare) Then
    'Userlist for chat
    On Error Resume Next
    StrNameList = StrNameList & Napresponse
    If InStr(1, StrNameList, Chr(153) & Chr(1), vbTextCompare) Then
        'if end of userlist char is present
        napparsed() = Split(StrNameList, Chr(152) & Chr(1))
            For stp = LBound(napparsed) To UBound(napparsed)
                strlen = Len(napparsed(stp)) - 1
                If strlen > 0 Then
                    napparsed2() = Split(Left(napparsed(stp), strlen), " ")
                    a = napparsed2(1) '& " sharing " & napparsed(2)
                    If UBound(napparsed2) > 2 Then
                        'determine line speed
                        Select Case napparsed2(3)
                            Case "0"
                                B = "unknown"
                            Case "1"
                                B = "14.4K"
                            Case "2"
                                B = "28.8K"
                            Case "3"
                                B = "33.6K"
                            Case "4"
                                B = "56.7K"
                            Case "5"
                                B = "64K ISDN"
                            Case "6"
                                B = "128K ISDN"
                            Case "7"
                                B = "Cable"
                            Case "8"
                                B = "DSL"
                            Case "9"
                                B = "T1"
                            Case "10"
                                B = "T3+"
                            Case Else
                                B = "unknown"
                        End Select
                    Else
                        B = "unknown"
                    End If
                    a = a & " on " & B
                    Frmchat!LstUserList.AddItem a
                End If
            Next
            StrNameList = ""
            Frmchat.Show
        End If

    ElseIf InStr(1, StrNameList, Chr(153) & Chr(1), vbTextCompare) Then
    'if end of user list char is present
    On Error Resume Next
    napparsed() = Split(StrNameList, Chr(152) & Chr(1))
            For stp = LBound(napparsed) To UBound(napparsed)
                strlen = Len(napparsed(stp)) - 1
                If strlen > 0 Then
                    napparsed2() = Split(Left(napparsed(stp), strlen), " ")
                    a = napparsed2(1) '& " sharing " & napparsed(2)
                    If UBound(napparsed2) > 2 Then
                        Select Case napparsed2(3)
                            Case "0"
                                B = "unknown"
                            Case "1"
                                B = "14.4K"
                            Case "2"
                                B = "28.8K"
                            Case "3"
                                B = "33.6K"
                            Case "4"
                                B = "56.7K"
                            Case "5"
                                B = "64K ISDN"
                            Case "6"
                                B = "128K ISDN"
                            Case "7"
                                B = "Cable"
                            Case "8"
                                B = "DSL"
                            Case "9"
                                B = "T1"
                            Case "10"
                                B = "T3+"
                            Case Else
                                B = "unknown"
                        End Select
                    Else
                        B = "unknown"
                    End If
                    a = a & " on " & B
                    Frmchat!LstUserList.AddItem a
                End If
            Next
            StrNameList = ""
            Frmchat.Show

    ElseIf InStr(1, Napresponse, Chr(149) & Chr(1), vbTextCompare) Or InStr(1, Napresponse, Chr(154) & Chr(1), vbTextCompare) Then
    'MsgBox afterit(Chr(1), Napresponse)
    ElseIf InStr(1, Napresponse, Chr(150) & Chr(1), vbTextCompare) Then
    'someone has joined the chat
        On Error Resume Next
        If Trim(beforeit(" ", afterit(Chr(1), Napresponse))) = Frmchat.Caption Then
            Napresponse = afterit(" ", Napresponse)
            napparsed() = Split(Napresponse, " ")

            Napresponse = "<" & napparsed(0) & " joined><sharing " & napparsed(1) & "><"
                If UBound(napparsed) > 1 Then
                    Select Case napparsed(2)
                        Case "0"
                            B = "unknown"
                        Case "1"
                            B = "14.4K"
                        Case "2"
                            B = "28.8K"
                        Case "3"
                            B = "33.6K"
                        Case "4"
                            B = "56.7K"
                        Case "5"
                            B = "64K ISDN"
                        Case "6"
                            B = "128K ISDN"
                        Case "7"
                            B = "Cable"
                        Case "8"
                            B = "DSL"
                        Case "9"
                            B = "T1"
                        Case "10"
                            B = "T3+"
                        Case Else
                            B = "unknown"
                    End Select
                Else
                    B = "unknown"
                End If
                Napresponse = Napresponse & B & ">"
                Frmchat!LstUserList.AddItem napparsed(0) & " on " & B
                Frmchat!TxtChat = Frmchat!TxtChat & vbCrLf & Napresponse
            Else
                Frmchat.Caption = Trim(beforeit(" ", afterit(Chr(1), Napresponse)))
            End If

    ElseIf InStr(1, Napresponse, Chr(151) & Chr(1), vbTextCompare) Then
    ' someone has left
    If Trim(beforeit(" ", afterit(Chr(1), Napresponse))) = Frmchat.Caption Then
        On Error Resume Next
        Napresponse = afterit(" ", Napresponse)
        napparsed() = Split(Napresponse, " ")
        'removes person from list

        For stp2 = 0 To Frmchat!LstUserList.ListCount - 1
            If beforeit(" ", Frmchat!LstUserList.List(stp2)) Like Trim(napparsed(0)) Then
            Frmchat!LstUserList.RemoveItem stp2
            End If
        Next stp2

        Napresponse = "<" & napparsed(0) & " left><sharing " & napparsed(1) & "><"
                If UBound(napparsed) > 1 Then
                    Select Case napparsed(2)
                        Case "0"
                            B = "unknown"
                        Case "1"
                            B = "14.4K"
                        Case "2"
                            B = "28.8K"
                        Case "3"
                            B = "33.6K"
                        Case "4"
                            B = "56.7K"
                        Case "5"
                            B = "64K ISDN"
                        Case "6"
                            B = "128K ISDN"
                        Case "7"
                            B = "Cable"
                        Case "8"
                            B = "DSL"
                        Case "9"
                            B = "T1"
                        Case "10"
                            B = "T3+"
                        Case Else
                            B = "unknown"
                    End Select
                Else
                B = "unknown"
                End If
                Napresponse = Napresponse & B & ">"
                Frmchat!TxtChat = Frmchat!TxtChat & vbCrLf & Napresponse
        Else
            Frmchat.Caption = Trim(beforeit(" ", afterit(Chr(1), Napresponse)))
        End If
        
    ElseIf InStr(1, Napresponse, Chr(145) & Chr(1), vbTextCompare) Then
    'you have parted the channel
    'MsgBox "You have left " & afterit(Chr(1), Napresponse)
    ElseIf InStr(1, Napresponse, Chr(240) & Chr(2), vbTextCompare) Then
        a = afterit(Chr(240) & Chr(2), Napresponse)
        Form2!Pingresponse.Text = a
    ElseIf (InStr(1, Napresponse, Chr(201), vbTextCompare) <> 0 And InStr(1, Napresponse, Chr(34), vbTextCompare) <> 0) Or InStr(1, Napresponse, Chr(202), vbTextCompare) <> 0 Then
        Form2!TxtResults = Form2!TxtResults & Napresponse
    Else
        Napresponse = Replace(Napresponse, Chr(0), "")
        MsgBox Napresponse
    End If
End Select
End Sub


Sub NapConnect()
If Step = 2 Then
    Txtinfo = Txtinfo & vbCrLf & "Sending login info..."
    ' as indicated in the napster protocol spec
    ' the format of each message to the server should be
    ' <2 byte length field><2 byte type field><data>
    ' where length specifies the length of the data

    Winsock1.SendData Chr(bytelen) & Chr(0) '2 byte length field
    Winsock1.SendData Chr(2) & Chr(0) '2 byte type field
    Winsock1.SendData Napinfo 'data field
End If
Text1.Enabled = False
End Sub

Sub MainConnect()
Winsock1.Close
While Winsock1.State
  DoEvents
Wend

Command1.Enabled = False
Command2.Enabled = True
Txtinfo = ""
Text4 = ""
Napinfo = Text1 & " " & Text2 & " " & Text3 & " " & Chr(34) & "v2.0 BETA 6" & Chr(34) & " " & CStr(Combo1.ListIndex)
bytelen = Len(Napinfo)
Txtinfo = "Connecting to main napster server..."
Winsock1.Connect "server.napster.com", "8875"
Step = 1
Timer1.Enabled = True
End Sub
Sub NapDisconnect()
Txtinfo = Txtinfo & vbCrLf & "Disconnecting from server."
Naplogon = False
Secnd = 0
Winsock1.Close
While Winsock1.State
  DoEvents
Wend
Timer1.Enabled = False
Command1.Enabled = True
Command2.Enabled = False
Text1.Enabled = True
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox Description, vbExclamation, "Client error!"
End Sub

