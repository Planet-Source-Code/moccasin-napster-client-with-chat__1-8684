Attribute VB_Name = "Module1"
Type MyLong
    Value As Long
    End Type


Type MyIP
    a As Byte
    B As Byte
    C As Byte
    d As Byte
    End Type
Public Function afterit(afterwhat As String, inwhat As String) As String
Dim d As Integer, a As Integer, f As String
If InStr(1, inwhat, afterwhat, vbTextCompare) <> 0 Then
d = InStr(1, inwhat, afterwhat, vbTextCompare)
a = d + Len(afterwhat)
f = Mid$(inwhat, a, 500)
f = Trim$(f)
If f <> "" Then
afterit = f
End If
End If
End Function
Public Function beforeit(beforewhat As String, inwhat As String) As String
Dim d As Integer, a As Integer, f As String
If InStr(1, inwhat, beforewhat, vbTextCompare) <> 0 Then
d = InStr(1, inwhat, beforewhat, vbTextCompare)
a = d - 1
f = Left$(inwhat, a)
f = Trim$(f)
If f <> "" Then
beforeit = f
End If
End If
End Function

Public Function inbetween(inwhat As String, afterwhat As String, beforewhat As String) As String
Dim likewhoa As String
likewhoa = beforeit(beforewhat, afterit(afterwhat, inwhat))
inbetween = likewhoa$
End Function
Public Sub Pause(Duration As Long)
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= Duration
        DoEvents
    Loop
End Sub

Function GetLast(inwhat As String, Delimiter As String) As String
    Dim point As Integer
    point = InStrRev(inwhat, Delimiter)


    If point <> 0 Then
        GetLast = Trim(Mid(inwhat, (point + Len(Delimiter)), (Len(inwhat) - point) + Len(Delimiter)))
    Else
        GetLast = inwhat
    End If
End Function
Function IPToString(Value As Double) As String
    Dim l As MyLong
    Dim i As MyIP
    l.Value = DoubleToLong(Value)
    LSet i = l
    IPToString = i.a & "." & i.B & "." & i.C & "." & i.d
End Function


Function DoubleToLong(Value As Double) As Long


    If Value <= 2147483647 Then


        DoubleToLong = Value
        Else


            DoubleToLong = -(4294967296# - Value)
            End If
        End Function

Public Sub Killdupes(ListBox As Object)
'deletes duplicates from list
    Dim FirstCount As Long, SecondCount As Long
    On Error Resume Next
    For FirstCount& = 0& To ListBox.ListCount - 1
        For SecondCount& = 0& To ListBox.ListCount - 1
            If LCase(ListBox.list(FirstCount&)) = LCase(ListBox.list(SecondCount&)) And FirstCount& <> SecondCount& Then
                ListBox.RemoveItem SecondCount&
            End If
        Next SecondCount&
    Next FirstCount&
End Sub
