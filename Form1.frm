VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "DNAS Patcher to PNACH"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6825
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
   ScaleHeight     =   4800
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   6360
      Top             =   4320
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   0
      Text            =   "Form1.frx":0006
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "DNAS Patcher by kHn"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4920
      TabIndex        =   5
      Top             =   4440
      Width           =   1605
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "www.VTS-Tech.org"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   1410
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Veritas Technical Solutions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      TabIndex        =   3
      Top             =   4200
      Width           =   1920
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Written by VTSTech"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   4200
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InputLenPrev, InputLen, LinesOut, LinesIn, Data
Public Function Parse_Line(Data)
tmp = Split(Data, " ")
If Not tmp(0) = "DNAS-net" And Not tmp(0) = "Scanning" Then
    If Len(tmp(0)) = 8 Then
        addr = tmp(0)
        patch = tmp(1)
        Parse_Line = "patch=1,EE," & addr & ",word," & patch
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
Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
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
