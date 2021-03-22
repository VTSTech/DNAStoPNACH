VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "DNAS Patcher to PNACH"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7830
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   5280
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Use Enable"
      Height          =   210
      Left            =   3720
      TabIndex        =   6
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   330
      Left            =   7080
      TabIndex        =   5
      Top             =   4080
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Left            =   7320
      Top             =   4440
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   3720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0006
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "DNAS Patcher by kHn"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   5760
      TabIndex        =   4
      Top             =   4680
      Width           =   1995
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "www.VTS-Tech.org"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   0
      TabIndex        =   3
      Top             =   4680
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Written by: VTSTech/Veritas Technical Solutions"
      Height          =   210
      Left            =   1440
      TabIndex        =   2
      Top             =   4440
      Width           =   4485
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InputLenPrev, InputLen, LinesOut, LinesIn, Data, Build
Public Function Parse_Line(Data)
tmp = Split(Data, " ")
If Not tmp(0) = "DNAS-net" And Not tmp(0) = "Scanning" Then
    If Len(tmp(0)) = 8 Then
        addr = tmp(0)
        patch = tmp(1)
        If Mid(tmp(0), 1, 1) = 9 And Check1.Value = 1 Then
            Parse_Line = "patch=1,EE," & addr & ",extended," & patch
        ElseIf Mid(tmp(0), 1, 1) = 9 And Check1.Value = 0 Then
            a = a
        Else
            Parse_Line = "patch=1,EE," & addr & ",extended," & patch
        End If
    End If
End If

'DNAS-net Patcher (2020/07/13, test 21)

'Scanning file...
'[MODE 1] sceDNAS2GetStatus was found at offset E537DBE8h

'Enable Code
'907F6508 0C1FD8EA
'Stat Poke:
'D1F6E680 00000007
'01F6E680 00000005
'D1F6E680 00000006
'01F6E680 00000005

'Error Code:
'D1F6E684 ????????
'21F6E684 ????????

'Fake Deinit:
'D1F5DF7C 00000001
'01F5DF7C 00000000

'Still scanning...
'[MODE 3] SetStatus 6 was found at offset E53A9DBCh

'Enable Code
'907F6508 0C1FD8EA
'Still scanning...
'Scan ended
End Function

Private Sub Command1_Click()
Set FSO = CreateObject("Scripting.FileSystemObject")
Close #1
fn = Text3.Text
fn = VB.App.Path & "\" & fn
Open fn For Output As #1
Write #1, Text2.Text
Close #1
MsgBox "File written to " & fn
End Sub

Private Sub Form_Load()
Build = "0.1-R5"
Form1.Caption = "DNAS Patcher to PNACH `v" & Build & " by VTSTech"
Text1.Text = ""
Text2.Text = ""
Text3.Text = "game.pnach"
Timer1.Interval = 2000
Timer1.Enabled = True
InputLenPrev = Len(Text1.Text)

End Sub

Private Sub Label3_Click()
Shell ("start https://www.VTS-Tech.org/")
End Sub

Private Sub Label4_Click()
Shell ("start https://www.psx-place.com/threads/dnas-net-patcher.22813/")
End Sub

Private Sub Timer1_Timer()
InputLen = Len(Text1.Text)
Text2.Text = ""
DoEvents
If InputLen > InputLenPrev And InputLen >= 1 Then
    LinesIn = Split(Text1.Text, vbCrLf)
    For x = 0 To UBound(LinesIn)
    LinesOut = LinesIn(x)
    If Len(LinesOut) >= 1 Then
        LinesOut = Parse_Line(LinesOut)
        If Len(LinesOut) >= 1 Then
            Text2.Text = Text2.Text & LinesOut & vbCrLf
        End If
    End If
Next x
End If
End Sub
