VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Encryption Project"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Help"
      Height          =   465
      Left            =   2730
      TabIndex        =   10
      Top             =   3900
      Width           =   1245
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   465
      Left            =   1410
      TabIndex        =   9
      Top             =   3900
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encrypt"
      Default         =   -1  'True
      Height          =   465
      Left            =   90
      TabIndex        =   8
      Top             =   3900
      Width           =   1245
   End
   Begin VB.Frame Frame4 
      Caption         =   "Encrypted ciphertext:"
      Height          =   1815
      Left            =   2490
      TabIndex        =   5
      Top             =   1980
      Width           =   3075
      Begin VB.TextBox Text2 
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   2865
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Enter messagetext:"
      Height          =   1815
      Left            =   2490
      TabIndex        =   4
      Top             =   90
      Width           =   3075
      Begin VB.TextBox Text1 
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "Form1.frx":0000
         Top             =   270
         Width           =   2865
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select type of ciphertext"
      Height          =   1815
      Left            =   90
      TabIndex        =   2
      Top             =   1980
      Width           =   2295
      Begin VB.ListBox List2 
         Height          =   1425
         ItemData        =   "Form1.frx":000C
         Left            =   120
         List            =   "Form1.frx":0025
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   270
         Width           =   2085
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select type of messagetext"
      Height          =   1815
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   2295
      Begin VB.ListBox List1 
         Height          =   1425
         ItemData        =   "Form1.frx":006D
         Left            =   120
         List            =   "Form1.frx":0086
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   270
         Width           =   2085
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public dif As Integer
Public keyword As String

'Ascii
'Backwards
'Binary
'Caesar
'Pig Latin
'Random
'Transposition

Private Sub Command1_Click()
If List1.Text = List2.Text Then Text2.Text = Text1.Text: Exit Sub
Dim e As New clsEncrypt, t As String
If List1.Text = "Ascii" Then
    Select Case List2.Text
        Case "Backwards": Text2.Text = e.Backwards(Text1.Text)
        Case "Binary": Text2.Text = e.Ascii_Binary(Text1.Text)
        Case "Caesar": Text2.Text = e.Ascii_Caesar(Text1.Text, dif)
        Case "Pig Latin": Text2.Text = e.Ascii_PigLatin(Text1.Text)
        Case "Random": Text2.Text = e.Ascii_Random(Text1.Text, keyword)
        Case "Transposition": Text2.Text = e.Ascii_Transposition(Text1.Text)
    End Select
Else
    Select Case List1.Text
        Case "Backwards": t = e.Backwards(Text1.Text)
        Case "Binary": t = e.Binary_Ascii(Text1.Text)
        Case "Caesar": t = e.Ascii_Caesar(Text1.Text, dif * -1)
        Case "Pig Latin": t = e.PigLatin_Ascii(Text1.Text)
        Case "Random": t = e.Random_Ascii(Text1.Text, keyword)
        Case "Transposition": t = e.Transposition_Ascii(Text1.Text)
    End Select
    If t = "" Then Text2.Text = "Error!": Exit Sub
    Select Case List2.Text
        Case "Ascii": Text2.Text = t
        Case "Backwards": Text2.Text = e.Backwards(t)
        Case "Binary": Text2.Text = e.Ascii_Binary(t)
        Case "Caesar": Text2.Text = e.Ascii_Caesar(t, dif)
        Case "Pig Latin": Text2.Text = e.Ascii_PigLatin(t)
        Case "Random": Text2.Text = e.Ascii_Random(t, keyword)
        Case "Transposition": Text2.Text = e.Ascii_Transposition(t)
    End Select
End If

End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Command3_Click()
MsgBox "How to use:" & vbCrLf & vbCrLf & _
    vbTab & "1. Choose what type of text is in the top textbox by selecting the appropriate type in the top listbox." & vbCrLf & _
    vbTab & vbTab & "Ex. If you entered binary code into the top textbox, choose Binary from the top listbox." & vbCrLf & _
    vbTab & "2. Choose what type of encryption you want to use to convert the text." & vbCrLf & vbCrLf & _
    "Tips:" & vbCrLf & vbCrLf & _
    vbTab & "- If you want to convert encrypted text back into its original form," & vbCrLf & "select Ascii in the bottom listbox and whatever type of encrypted text it is in the top listbox." & _
    "", , "Help"
End Sub

Private Sub Form_Load()
List1.Selected(0) = True
List2.Selected(0) = True
End Sub

Private Sub List1_Click()
If List1.Text = "Caesar" Then dif = Val(InputBox("Enter difference:"))
If List1.Text = "Random" Then keyword = InputBox("Enter keyword:")
End Sub

Private Sub List2_Click()
If List2.Text = "Caesar" Then dif = Val(InputBox("Enter difference:"))
If List2.Text = "Random" Then keyword = InputBox("Enter keyword:")
End Sub
