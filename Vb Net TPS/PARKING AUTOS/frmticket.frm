VERSION 5.00
Begin VB.Form frmticket 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ticket"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdRecalcular 
      Caption         =   "Recalcular"
      Height          =   495
      Left            =   2160
      TabIndex        =   10
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox lblimporte 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox lblduracion 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox lblhfinalizacion 
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox lblhinicio 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox lblpatente 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Importe a pagar:"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Duración:"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Hora de finalizacion:"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Hora de inicio:"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Patente:"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frmticket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCerrar_Click()

frmticket.Hide

End Sub

Private Sub cmdRecalcular_Click()

intervalo = Format(TimeValue(lblhfinalizacion.Text) - TimeValue(lblhinicio.Text), "hh:mm:ss")
mins = Mid(intervalo, 4, 1)
horas = Left(intervalo, 2)

If horas = 1 Or horas = 0 Then
    precio = 40
End If

If horas = 2 Then
    precio = 80
End If

If horas = 3 Then
    precio = 120
End If

If horas > 0 And mins > 0 And mins < 15 Then
    precio = precio + 10
End If

If horas > 0 And mins > 15 And mins < 30 Then
    precio = precio + 20
End If

If horas > 0 And mins > 30 And mins < 45 Then
    precio = precio + 30
End If

If horas > 0 And mins > 30 And mins < 45 Then
    precio = precio + 40
End If

If horas > 3 Then
    precio = 130
End If


lblduracion.Text = intervalo
lblimporte.Text = "$ " & precio

End Sub

Private Sub Form_Activate()

If lblduracion = "Estadia" Then
    lblhinicio.Enabled = False
    lblhfinalizacion.Enabled = False
    cmdRecalcular.Enabled = False
Else
    lblhinicio.Enabled = True
    lblhfinalizacion.Enabled = True
    cmdRecalcular.Enabled = True
End If

End Sub

