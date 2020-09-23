VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Encryption Module"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   735
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   2400
      Width           =   3255
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Standard With Alpha Decoding"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   1920
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Standard Decryption"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Standard With Alpha Encoding"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Standard Encryption"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Text            =   "String to encrypt"
      Top             =   0
      Width           =   6855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   6960
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   3480
      X2              =   3480
      Y1              =   360
      Y2              =   3240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub Command1_Click()
    Text2 = Encrypt(Text1)
End Sub

Private Sub Command2_Click()
    Text3 = Encrypt(Text1, True)
End Sub

Private Sub Command3_Click()
    Text4 = Decrypt(Text2)
End Sub

Private Sub Command4_Click()
    Text5 = Decrypt(Text3, True)
End Sub
