VERSION 5.00
Begin VB.Form frmFracasado 
   Caption         =   "Fracasado"
   ClientHeight    =   1575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblFracasado 
      Caption         =   "Usted no ha sido capaz de adivinar el número."
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "frmFracasado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSalir_Click()
End
End Sub

