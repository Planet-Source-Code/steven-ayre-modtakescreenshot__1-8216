VERSION 5.00
Begin VB.Form frmScreenShot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Example of how to use modTakeScreenShot"
   ClientHeight    =   5052
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   5376
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5052
   ScaleWidth      =   5376
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPartial 
      Caption         =   "Take a &Partial screenshot instead"
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   5172
   End
   Begin VB.CommandButton cmdScreenShot 
      Caption         =   "&Click here to take a screenshot"
      Default         =   -1  'True
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   5172
   End
   Begin VB.PictureBox picScreenShot 
      Height          =   3840
      Left            =   120
      ScaleHeight     =   3792
      ScaleWidth      =   5076
      TabIndex        =   1
      Top             =   120
      Width           =   5120
   End
End
Attribute VB_Name = "frmScreenShot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPartial_Click()
frmPartialScreenShot.Show
frmScreenShot.Hide
End Sub

Private Sub cmdScreenShot_Click()
Call ScreenShot(picScreenShot.hDC)
End Sub
