VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEncrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim intBin(7) As Integer

'http://www.achiever.com/freehmpg/cryptology/cryptofr.html
'http://www-theory.dcs.st-and.ac.uk/~sal/school/CS3010/Lectures/forhtml/node4.html

Private Sub Class_Initialize()
intBin(0) = 1
intBin(1) = 2
intBin(2) = 4
intBin(3) = 8
intBin(4) = 16
intBin(5) = 32
intBin(6) = 64
intBin(7) = 128
End Sub

Public Function Ascii_Binary(str As String) As String
Dim i As Integer, h As Integer, t As Integer, s As String, m As Integer
For m = 1 To Len(str)
    i = Asc(Mid(str, m, 1))
    For h = 7 To 0 Step -1
        If intBin(h) <= i - t Then
            s = s & "1"
            t = t + intBin(h)
        Else
            s = s & "0"
        End If
    Next h
    t = 0
Next m
Ascii_Binary = Trim(s)
End Function

Public Function Binary_Ascii(bin As String) As String
Dim i As Integer, t As Integer, m As Integer, s As String, n As Integer
For m = 1 To Len(bin) / 8
    For i = ((m - 1) * 8) + 1 To ((m - 1) * 8) + 8
        n = n + 1
        t = t + (Val(Mid(bin, i, 1)) * intBin(8 - n))
    Next i
    s = s & Chr(t)
    t = 0: n = 0
Next m
Binary_Ascii = s
End Function

Public Function Backwards(str As String) As String
Dim i As Integer, s As String
For i = Len(str) To 1 Step -1
    s = s & Mid(str, i, 1)
Next i
Backwards = s
End Function

Public Function Ascii_Caesar(str As String, dif As Integer) As String
checkdif:
If dif > 26 Then
    dif = dif - 26
    GoTo checkdif
End If
Dim i As Integer, s As String, t As String
For i = 1 To Len(str)
    t = Mid(str, i, 1)
    If t = " " Then s = s & " ": GoTo WasSpace
    If t = LCase(t) Then 'lowercase: 97-122
        If Asc(t) + dif > 122 Then
            s = s & Chr(97 + ((Asc(t) + dif) - 122))
        Else
            s = s & Chr(Asc(t) + dif)
        End If
    Else 'uppercase: 65-90
        If Asc(t) + dif > 90 Then
            s = s & Chr(65 + ((Asc(t) + dif) - 90))
        Else
            s = s & Chr(Asc(t) + dif)
        End If
    End If
WasSpace:
Next i
Ascii_Caesar = s
End Function

Public Function Ascii_Transposition(str As String) As String
Dim i As Integer, s As String, d As Integer
d = Format(Len(str) / 2, 0)
For i = 1 To Len(str) / 2
    s = s & Mid(str, i, 1)
    s = s & Mid(str, i + d, 1)
Next i
If Len(str) / 2 < d Then s = s & Mid(str, d, 1)
Ascii_Transposition = s
End Function

Public Function Transposition_Ascii(str As String) As String
Dim i As Integer, s As String, d As Integer
d = Len(str) / 2
For i = 1 To Len(str) Step 2
    s = s & Mid(str, i, 1)
Next i
For i = 2 To Len(str) Step 2
    s = s & Mid(str, i, 1)
Next i
Transposition_Ascii = s
End Function

Public Function Ascii_PigLatin(str As String) As String
Dim i As Integer, s As String, n() As String, c As Integer
s = str
Do: DoEvents
    i = InStr(s, " ")
    If i = 0 Then Exit Do
    s = Mid(s, i + 1)
    c = c + 1
Loop
n = Split(str, " ")
s = ""
For i = 0 To c
    s = s & Right(n(i), Len(n(i)) - 1) & Left(n(i), 1) & "ay "
Next i
s = Trim(s)
Ascii_PigLatin = s
End Function

Public Function PigLatin_Ascii(str As String) As String
Dim i As Integer, s As String, n() As String, c As Integer
s = str
Do: DoEvents
    i = InStr(s, " ")
    If i = 0 Then Exit Do
    s = Mid(s, i + 1)
    c = c + 1
Loop
n = Split(str, " ")
s = ""
For i = 0 To c
    n(i) = Left(n(i), Len(n(i)) - 2)
    s = s & Right(n(i), 1) & Left(n(i), Len(n(i)) - 1) & " "
Next i
s = Trim(s)
PigLatin_Ascii = s
End Function

Function Ascii_Random(CustomValue As String, EncryptCode As String) As String
Dim a$, i As Long, b$, d$, q%, c As String
q = Val(Trigger(EncryptCode))
a = CustomValue
For i = 1 To Len(a)
    DoEvents
    b = Mid(a, i, 1)
    c = Asc(b) + q
    If c < 0 Then c = c - c - c
    If Len(c) = 1 Then c = "00" & c
    If Len(c) = 2 Then c = "0" & c
    d = d & c
Next i
Dim f As Integer, e As String, g As Integer
For i = 1 To Len(d)
    DoEvents
    f = Mid(d, i, 1)
    Randomize
    g = Rnd * 19
    Select Case f
        Case 0: e = e & Chr(g + 55)
        Case 1: e = e & Chr(g + 75)
        Case 2: e = e & Chr(g + 95)
        Case 3: e = e & Chr(g + 115)
        Case 4: e = e & Chr(g + 135)
        Case 5: e = e & Chr(g + 155)
        Case 6: e = e & Chr(g + 175)
        Case 7: e = e & Chr(g + 195)
        Case 8: e = e & Chr(g + 215)
        Case 9: e = e & Chr(g + 235)
    End Select
Next i
Ascii_Random = e
End Function

Function Random_Ascii(ByVal CustomValue As String, ByVal EncryptCode As String) As String
Dim X As String, Y As String, z As String, q As Integer
X = Left(EncryptCode, 1)
Y = Right(EncryptCode, 1)
z = Mid(EncryptCode, Format(Len(EncryptCode) / 2, "#"), 1)
q = Format((Asc(X) + Asc(Y) + Asc(z)) / 6, "##")
Dim a As String, i As Long, b As String, c As String
a = CustomValue
For i = 1 To Len(a)
    DoEvents
    b = Val(Asc(Mid(a, i, 1)))
    If b >= 55 And b < 75 Then c = c & 0
    If b >= 75 And b < 95 Then c = c & 1
    If b >= 95 And b < 115 Then c = c & 2
    If b >= 115 And b < 135 Then c = c & 3
    If b >= 135 And b < 155 Then c = c & 4
    If b >= 155 And b < 175 Then c = c & 5
    If b >= 175 And b < 195 Then c = c & 6
    If b >= 195 And b < 215 Then c = c & 7
    If b >= 215 And b < 235 Then c = c & 8
    If b >= 235 And b < 255 Then c = c & 9
    If b < 55 Or b >= 255 Then
        Random_Ascii = ""
        Exit Function
    End If
Next i
Dim d As Integer, e As String, f As Integer
For i = 1 To Len(c) Step 3
    DoEvents
    d = Mid(c, i, 3)
    f = d - q
    If f < 0 Then f = f - f - f
    If f > 255 Or f < 0 Then
        Random_Ascii = ""
        Exit Function
    End If
    e = e & Chr(f)
Next i
Random_Ascii = e
End Function

Private Function Trigger(CodeWord As String) As String
If CodeWord = "" Then Trigger = "": Exit Function
Dim X$, Y$, z$, q%
X = Left(CodeWord, 1)
Y = Right(CodeWord, 1)
z = Mid(CodeWord, Format(Len(CodeWord) / 2, "#"), 1)
q = Format((Asc(X) + Asc(Y) + Asc(z)) / 6, "##")
Trigger = q
End Function
