VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Numbers To Numerals"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7275
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2085
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   855
      Left            =   6120
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Data"
      Height          =   855
      Left            =   6120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtEnglish 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label lblDisplayExtra 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   5895
   End
   Begin VB.Label lblDisplayExtra 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   5895
   End
   Begin VB.Label lblDisplay 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lCoin(1 To 25) As Long
Dim sRome(1 To 25) As String
Dim lValue As Long
Dim i As Integer
Dim sRoman As String
Dim sExtra As String
Dim sUnder As String

Private Sub txtEnglish_Change()
'
lblDisplay = ""
lblDisplayExtra(0) = ""
lblDisplayExtra(1) = ""
sExtra = ""
sRoman = ""
lValue = Val(txtEnglish)
'
If lValue > 0 Then
    For i = 25 To 1 Step -1
        Do
            If lValue >= lCoin(i) Then
                lValue = lValue - lCoin(i)
                sRoman = sRoman + sRome(i)
                If i > 13 Then
                    sExtra = sExtra + sUnder
                    Else
                    If i = 1 Or i = 3 Or i = 5 Or i = 7 Or i = 9 Or i = 11 Or i = 13 Or i = 15 Or i = 17 Or i = 19 Or i = 21 Or i = 23 Or i = 25 Then
                        sExtra = sExtra + " "
                        Else
                        sExtra = sExtra + "  "
                    End If
                End If
            End If
        Loop Until lValue < lCoin(i)
    Next i
End If
'
'sExtra = sRoman
lblDisplay = sRoman
lblDisplayExtra(0) = sExtra
lblDisplayExtra(1) = sExtra
'
End Sub

Private Sub cmdClear_Click()
txtEnglish = ""
lblDisplay = ""
lblDisplayExtra(0) = ""
lblDisplayExtra(1) = ""
sRoman = ""
sExtra = ""
txtEnglish.SetFocus
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub Form_Load()
'
lCoin(1) = 1: sRome(1) = "I"
lCoin(2) = 4: sRome(2) = "IV"
lCoin(3) = 5: sRome(3) = "V"
lCoin(4) = 9: sRome(4) = "IX"
lCoin(5) = 10: sRome(5) = "X"
lCoin(6) = 40: sRome(6) = "XL"
lCoin(7) = 50: sRome(7) = "L"
lCoin(8) = 90: sRome(8) = "XC"
lCoin(9) = 100: sRome(9) = "C"
lCoin(10) = 400: sRome(10) = "CD"
lCoin(11) = 500: sRome(11) = "D"
lCoin(12) = 900: sRome(12) = "CM"
lCoin(13) = 1000: sRome(13) = "M"
lCoin(14) = 4000: sRome(14) = "MV"
lCoin(15) = 5000: sRome(15) = "V"
lCoin(16) = 9000: sRome(16) = "MX"
lCoin(17) = 10000: sRome(17) = "X"
lCoin(18) = 40000: sRome(18) = "XL"
lCoin(19) = 50000: sRome(19) = "L"
lCoin(20) = 90000: sRome(20) = "XC"
lCoin(21) = 100000: sRome(21) = "C"
lCoin(22) = 400000: sRome(22) = "CD"
lCoin(23) = 500000: sRome(23) = "D"
lCoin(24) = 900000: sRome(24) = "CM"
lCoin(25) = 1000000: sRome(25) = "M"
'
sUnder = "_"
End Sub

