VERSION 5.00
Begin VB.Form frmPuntuacion 
   Caption         =   "Form1"
   ClientHeight    =   6045
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14505
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   14505
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir2 
      Caption         =   "Salir"
      Height          =   615
      Left            =   5760
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdGame2 
      Caption         =   "Volver Al Juego"
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.ListBox listpuntuacion 
      Height          =   2010
      Left            =   7800
      TabIndex        =   1
      Top             =   3360
      Width           =   4455
   End
   Begin VB.ListBox listplayer 
      Height          =   2010
      Left            =   1080
      TabIndex        =   0
      Top             =   3360
      Width           =   4455
   End
End
Attribute VB_Name = "frmPuntuacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGame2_Click()
contador2 = 0
frmGame.Show
frmPuntuacion.Hide

End Sub

Private Sub cmdSalir2_Click()
MsgBox "Desea salir del programa?", vbYesNo, "Salir"
If vbYes Then
End
End If
End Sub

Private Sub Form_Activate()
listplayer.AddItem nombre
listpuntuacion.AddItem contador2

For i = 0 To listpuntuacion.ListCount - 2
    For J = i + 1 To listpuntuacion.ListCount - 1
        If Val(listpuntuacion.List(i)) > Val(listpuntuacion.List(J)) Then
            aux = listpuntuacion.List(i)
            listpuntuacion.List(i) = listpuntuacion.List(J)
            listpuntuacion.List(J) = aux
            aux = listplayer.List(i)
            listplayer.List(i) = listplayer.List(J)
            listplayer.List(J) = aux
        End If
    Next J
Next i
End Sub
