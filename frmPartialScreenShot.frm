VERSION 5.00
Begin VB.Form frmPartialScreenShot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Example of how to use modTakeScreenShot"
   ClientHeight    =   5772
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   5388
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5772
   ScaleWidth      =   5388
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFull 
      Caption         =   "Take a &Full screenshot instead"
      Height          =   372
      Left            =   120
      TabIndex        =   11
      Top             =   5280
      Width           =   5172
   End
   Begin VB.CommandButton cmdScreenShot 
      Caption         =   "&Click here to take a screenshot"
      Default         =   -1  'True
      Height          =   372
      Left            =   120
      TabIndex        =   8
      Top             =   4800
      Width           =   5172
   End
   Begin VB.PictureBox picScreenShot 
      Height          =   3840
      Left            =   120
      ScaleHeight     =   3792
      ScaleWidth      =   5076
      TabIndex        =   10
      Top             =   840
      Width           =   5120
   End
   Begin VB.TextBox txtHeight 
      Height          =   288
      Left            =   2640
      TabIndex        =   7
      Top             =   480
      Width           =   732
   End
   Begin VB.TextBox txtWidth 
      Height          =   288
      Left            =   2640
      TabIndex        =   5
      Top             =   120
      Width           =   732
   End
   Begin VB.TextBox txtLeft 
      Height          =   288
      Left            =   840
      TabIndex        =   3
      Top             =   480
      Width           =   732
   End
   Begin VB.TextBox txtTop 
      Height          =   288
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   732
   End
   Begin VB.Label lblNote 
      Caption         =   "N.B. These values are in pixels."
      Height          =   492
      Left            =   3600
      TabIndex        =   9
      Top             =   240
      Width           =   1572
   End
   Begin VB.Label lblTop 
      Caption         =   "&Top"
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   612
   End
   Begin VB.Label lblWidth 
      Caption         =   "&Width"
      Height          =   252
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   612
   End
   Begin VB.Label lblHeight 
      Caption         =   "&Height"
      Height          =   252
      Left            =   1920
      TabIndex        =   6
      Top             =   480
      Width           =   612
   End
   Begin VB.Label lblLeft 
      Caption         =   "&Left"
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   612
   End
End
Attribute VB_Name = "frmPartialScreenShot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFull_Click()
frmScreenShot.Show
frmPartialScreenShot.Hide
End Sub

Private Sub cmdScreenShot_Click()
On Error GoTo Err 'in case the .Text are not numeric
Dim Top As Long, Left As Long, Width As Long, Height As Long
Top = CLng(txtTop.Text)
Left = CLng(txtLeft.Text)
Width = CLng(txtWidth.Text)
Height = CLng(txtHeight.Text)
Call PartialScreenShot(picScreenShot.hDC, Top, Left, Width, Height)
Err:
End Sub

Private Sub txtTop_GotFocus()
txtTop.SelStart = 0
txtTop.SelLength = Len(txtTop.Text)
End Sub

Private Sub txtLeft_GotFocus()
txtLeft.SelStart = 0
txtLeft.SelLength = Len(txtLeft.Text)
End Sub

Private Sub txtWidth_GotFocus()
txtWidth.SelStart = 0
txtWidth.SelLength = Len(txtWidth.Text)
End Sub

Private Sub txtHeight_GotFocus()
txtHeight.SelStart = 0
txtHeight.SelLength = Len(txtHeight.Text)
End Sub

