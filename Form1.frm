VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Relevancey 2"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Pic2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   2280
      Index           =   2
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   10
      Tag             =   "Empty1"
      Top             =   4680
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.PictureBox Pic2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   2280
      Index           =   0
      Left            =   2760
      Picture         =   "Form1.frx":1091A
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   1
      Tag             =   "A"
      Top             =   360
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.PictureBox Pic2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   2280
      Index           =   1
      Left            =   2520
      Picture         =   "Form1.frx":21234
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   9
      Tag             =   "B"
      Top             =   360
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear Box"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   2295
   End
   Begin VB.PictureBox picBlack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   7440
      Picture         =   "Form1.frx":31B4E
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   6
      Top             =   6240
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox picWhite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   7320
      Picture         =   "Form1.frx":31BE0
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   5
      Top             =   6240
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   2280
      Left            =   120
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   0
      Top             =   360
      Width           =   2280
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check Relevancey"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":31C72
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   4200
      Width           =   7335
   End
   Begin VB.Label Pcnt 
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   2760
      Width           =   4815
   End
   Begin VB.Label Label3 
      Caption         =   "Mask"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Drawn"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Arr1(0 To 29, 0 To 29) As Integer
Dim Arr2(0 To 29, 0 To 29) As Integer
Dim Relevance() As Integer
Dim Rel As Integer
Dim Alphabet As Integer

Private Sub Command1_Click()
On Local Error Resume Next
Dim X2, Y2 As Integer
Dim Intloop, Curvar, AtPart, Highest As Integer

Alphabet = Pic2.UBound
ReDim Relevance(0 To Pic2.UBound)
Rel = -1
AtPart = 0
Highest = 0

For X = 0 To 149
    For Y = 0 To 149
        X2 = Int(X / 5)
        Y2 = Int(Y / 5)

        If GetPixel(pic1.hdc, X2 * 5, Y2 * 5) = vbWhite Then
            Arr1(X2, Y2) = 1
        Else
            Arr1(X2, Y2) = 0
        End If
    Next Y
Next X

'when there is a loop here for difrant images add a progress bar in it
For a = 0 To Alphabet

For X = 0 To 149
    For Y = 0 To 149
        X2 = Int(X / 5)
        Y2 = Int(Y / 5)

        If GetPixel(Pic2(a).hdc, X2 * 5, Y2 * 5) = vbWhite Then
            Arr2(X2, Y2) = 1
        Else
            Arr2(X2, Y2) = 0
        End If
    Next Y
Next X

Call CalculatePercent
Next a


For B = 0 To Alphabet
Pic2(B).Visible = False
If Relevance(B) > Curvar Then
Curvar = Relevance(B)
AtPart = B
Highest = B
End If
Next B

Pic2(Highest).Visible = True
Pcnt.Caption = "Relevance = " & Relevance(Highest) & "% Letter = " & Pic2(Highest).Tag

End Sub

Private Sub Command2_Click()
pic1.Cls
pic1.Picture = LoadPicture("")
End Sub

Private Sub pic1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Coloura As Long
Dim i, j As Integer
Coloura = RGB(255, 255, 255)

If Button = 1 Then
    For i = X - 4 To X + 4
        For j = Y - 4 To Y + 4
            SetPixel pic1.hdc, i, j, Coloura
        Next j
    Next i
    pic1.Refresh
End If
End Sub

Private Sub pic1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    pic1_MouseDown Button, Shift, X, Y
End If
End Sub

Sub CalculatePercent()
On Error Resume Next
Dim total As Double
Dim Cor As Double
Dim Cor2 As Double
Dim X2, Y2, B As Integer
total = 0

For X = 0 To 99
For Y = 0 To 99
X2 = Int(X / 5)
Y2 = Int(Y / 5)
If Arr2(X2, Y2) = 1 Then
  Cor2 = Cor2 + 1
End If
Next Y
Next X

For X = 0 To 99
For Y = 0 To 99
X2 = Int(X / 5)
Y2 = Int(Y / 5)
If Arr1(X2, Y2) = 1 And Arr2(X2, Y2) = 1 Then
  Cor = Cor + 1
End If
Next Y
Next X

total = (Cor / Cor2) * 100
total = Int(total)

Rel = Rel + 1
Relevance(Rel) = total


End Sub

