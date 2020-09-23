VERSION 5.00
Begin VB.Form Frmchatlist 
   BackColor       =   &H00FF8080&
   Caption         =   "Channel List"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2745
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   2745
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   1560
   End
   Begin VB.Frame Frame1 
      Caption         =   "Channel List"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.ListBox List1 
         Height          =   2790
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Channel                 # of people"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
   End
End
Attribute VB_Name = "Frmchatlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmr As Long
Private Sub Form_Load()
Unload Frmchat
End Sub

Private Sub List1_DblClick()
Dim strchannel As String, bytelen As String
If List1.ListIndex > -1 Then
strchannel = Trim(List1.list(List1.ListIndex))

Frmchat!LstUserList.Clear
        If InStr(1, strchannel, Chr(9), vbTextCompare) <> 0 Then
        strchannel = beforeit(Chr(9), strchannel)
        End If

Frmchat.Caption = strchannel
Form1!Txtinfo = Form1!Txtinfo & vbCrLf & "joining channel " & Chr(34) & strchannel & Chr(34)
bytelen = CStr(Len(strchannel))
Form1!Winsock1.SendData Chr(bytelen) & Chr(0)
Form1!Winsock1.SendData Chr(144) & Chr(1)
Form1!Winsock1.SendData strchannel
Frame1.Enabled = False
Timer1.Enabled = True
Else
MsgBox "Choose a channel to join!", vbExclamation, "Channel error"
End If
End Sub

Private Sub Timer1_Timer()
Dim strchannel As String, bytelen As String, userAction

If tmr < 10 Then
' if you have not joined a channel within 10 seconds then
tmr = tmr + 1
Else
userAction = MsgBox("Channel join timed out", vbRetryCancel, "Channel error")
    If userAction = vbRetry Then
        If List1.ListIndex > -1 Then
        Frmchat!LstUserList.Clear
        strchannel = Trim(List1.list(List1.ListIndex))
        If InStr(1, strchannel, Chr(9), vbTextCompare) <> 0 Then
        strchannel = beforeit(Chr(9), strchannel)
        End If
        
        Frmchat.Caption = strchannel

        bytelen = CStr(Len(strchannel))
        ' leave channel incase the server thinks you're in
        Form1!Winsock1.SendData Chr(bytelen) & Chr(0)
        Form1!Winsock1.SendData Chr(145) & Chr(1)
        Form1!Winsock1.SendData strchannel
        
        Frmchat!LstUserList.Clear
        
        ' join channel
        Form1!Winsock1.SendData Chr(bytelen) & Chr(0)
        Form1!Winsock1.SendData Chr(144) & Chr(1)
        Form1!Winsock1.SendData strchannel
        tmr = 0
        End If
    Else
        tmr = 0
        Frame1.Enabled = True
        Timer1.Enabled = False
        ' leave channel incase the server thinks you're in
        Form1!Winsock1.SendData Chr(bytelen) & Chr(0)
        Form1!Winsock1.SendData Chr(145) & Chr(1)
        Form1!Winsock1.SendData strchannel
        
    End If
End If

End Sub
