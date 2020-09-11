VERSION 5.00
Begin VB.Form frmMensaje 
   Caption         =   "Form1"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsalir 
      Caption         =   "salir"
      Height          =   735
      Left            =   2040
      TabIndex        =   3
      Top             =   3840
      Width           =   3975
   End
   Begin VB.CommandButton cmdborrar 
      Caption         =   "borrar"
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   3120
      Width           =   3735
   End
   Begin VB.CommandButton cmdescribir 
      Caption         =   "escribir"
      Height          =   555
      Left            =   2160
      TabIndex        =   1
      Top             =   2280
      Width           =   3855
   End
   Begin VB.TextBox txtmensaje 
      Height          =   975
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   5415
   End
End
Attribute VB_Name = "frmMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdborrar_Click()
txtmensaje.Text = ""

End Sub

Private Sub cmdescribir_Click()
txtmensaje.Text = "hola 2AO"
End Sub

Private Sub cmdsalir_Click()
End
End Sub
