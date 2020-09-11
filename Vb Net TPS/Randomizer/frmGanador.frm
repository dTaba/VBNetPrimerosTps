VERSION 5.00
Begin VB.Form frmGanador 
   Caption         =   "Ganador"
   ClientHeight    =   2640
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   4065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Puntuaciones"
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdJugarDeNuevo 
      Caption         =   "Jugar de nuevo"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lblGanador 
      Caption         =   "Label2"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.Image imgGanador 
      Height          =   2760
      Left            =   0
      Picture         =   "frmGanador.frx":0000
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   135
      Left            =   2400
      TabIndex        =   0
      Top             =   480
      Width           =   15
   End
End
Attribute VB_Name = "frmGanador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdJugarDeNuevo_Click()
frmGanador.Hide
frmGame.Show
contador2 = 0
End Sub

Private Sub cmdSalir_Click()
MsgBox "Desea salir del programa?", vbYesNo, "Salir"
If vbYes Then
End
End If
End Sub

Private Sub Command1_Click()
frmPuntuacion.Show
frmGanador.Hide
End Sub

Private Sub Form_Activate()
lblGanador.Caption = nombre & ", usted ha ganado en " & contador2 & " intentos!"
End Sub

