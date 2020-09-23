VERSION 5.00
Begin VB.Form Frmchat 
   BackColor       =   &H00FF8080&
   Caption         =   "Chat"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8865
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      MaxLength       =   200
      TabIndex        =   0
      Top             =   5160
      Width           =   8895
   End
   Begin VB.Frame Frame2 
      Caption         =   "Chat"
      Height          =   5175
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6375
      Begin VB.TextBox TxtChat 
         Height          =   4815
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "User list"
      Height          =   5175
      Left            =   6360
      TabIndex        =   1
      Top             =   0
      Width           =   2535
      Begin VB.CommandButton Command1 
         Caption         =   "Clear chat screen"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   4560
         Width           =   2175
      End
      Begin VB.ListBox LstUserList 
         Height          =   3960
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "# people"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   4320
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Frmchat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
TxtChat = ""
End Sub

' "whois" or "finger" a user cmd button
'Private Sub Command2_Click()
'Dim bytelen As String, strUser As String
'If LstUserList.ListIndex > 0 Then
'strUser = Trim(beforeit(" ", LstUserList.List(LstUserList.ListIndex)))
'bytelen = Len(strUser)

'If Form1!Winsock1.State = sckConnected And strUser <> "" Then
'Form1!Winsock1.SendData Chr(bytelen) & Chr(0)
'Form1!Winsock1.SendData Chr(91) & Chr(2)
'Form1!Winsock1.SendData strUser
'End If
'End If
'End Sub

Private Sub Form_Activate()
Unload Frmchatlist
Killdupes LstUserList
Text1 = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim bytelen As String
bytelen = CStr(Len(Frmchat.Caption))
Form1!Txtinfo = Form1!Txtinfo & vbCrLf & "leaving chat"
Form1!Winsock1.SendData Chr(bytelen) & Chr(0)
Form1!Winsock1.SendData Chr(145) & Chr(1)
Form1!Winsock1.SendData Frmchat.Caption
End Sub





Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim StrText As String, strchannel As String, Strdata As String, bytelen As String
If KeyAscii = 13 Then
StrText = Trim(Text1)
strchannel = Frmchat.Caption
Strdata = strchannel & " " & StrText
bytelen = CStr(Len(Strdata))
If Form1!Winsock1.State = sckConnected Then
Form1!Winsock1.SendData Chr(bytelen) & Chr(0)
Form1!Winsock1.SendData Chr(146) & Chr(1)
Form1!Winsock1.SendData Strdata
Text1 = ""
Else
MsgBox "Not connected.", vbExclamation, "Chat error"
End If
End If
End Sub

Private Sub TxtChat_Change()
On Error Resume Next
TxtChat.SelLength = 0
If Len(TxtChat.Text) > 0 Then
If Right$(TxtChat.Text, 1) = vbCrLf Then
TxtChat.SelStart = Len(TxtChat.Text) - 1
Exit Sub
End If
TxtChat.SelStart = Len(TxtChat.Text)
End If
Label1.Caption = CStr(LstUserList.ListCount) & " people in chat"
End Sub
