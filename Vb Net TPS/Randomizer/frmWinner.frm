VERSION 5.00
Begin VB.Form frmWinner 
   BackColor       =   &H00FFFFFF&
   Caption         =   "WINNER"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdJuego 
      Caption         =   "Regresar al juego"
      Height          =   615
      Left            =   3360
      TabIndex        =   3
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdJdN 
      Caption         =   "Puntuacion"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   4680
      Width           =   3015
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   1
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "¡Felicitaciones, usted ha adivinado el número                           en solo un intento!"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   3720
      Width           =   6255
   End
   Begin VB.Image imgWinner 
      Height          =   3600
      Left            =   0
      Picture         =   "frmWinner.frx":0000
      Top             =   0
      Width           =   7170
   End
End
Attribute VB_Name = "frmWinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdJdN_Click()
frmPuntuacion.Show
frmWinner.Hide
End Sub

Private Sub cmdJuego_Click()
frmGame.Show
frmWinner.Hide
contador2 = 0
End Sub

Private Sub cmdSalir_Click()
MsgBox "Desea salir del programa?", vbYesNo, "Salir"
If vbYes Then
End
End If
End Sub
