VERSION 5.00
Begin VB.Form frmTicket 
   Caption         =   "Ticket - CACHO TRANSPORTES"
   ClientHeight    =   6195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4500
   LinkTopic       =   "Form2"
   ScaleHeight     =   6195
   ScaleWidth      =   4500
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHI 
      Height          =   285
      Left            =   2280
      TabIndex        =   16
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtHF 
      Height          =   285
      Left            =   2280
      TabIndex        =   15
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cbRecalcular 
      Caption         =   "Recalcular"
      Height          =   495
      Left            =   3000
      TabIndex        =   10
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cbCerrar 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   1440
      TabIndex        =   9
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblImp 
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lblCarga 
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblDom 
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblTiempo 
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblGracias 
      Caption         =   "Gracias por Elegirnos!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label lblDominio 
      Caption         =   "Dominio:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblHoraInicio 
      Caption         =   "Hora de Inicio:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblHoraFinalización 
      Caption         =   "Hora de Finalización:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label lblDuracion 
      Caption         =   "Duración:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblImporte 
      Caption         =   " Importe a Abonar:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label lblCaragaTransportada 
      Caption         =   "Carga Transportada:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbCerrar_Click()

frmTicket.Hide

End Sub

Private Sub cbRecalcular_Click()

tiempo = Format(TimeValue(txtHF.Text) - TimeValue(txtHI.Text), "hh:mm")
lblTiempo.Caption = tiempo
                   
hora = Left(lblTiempo.Caption, 2)
min = Right(lblTiempo.Caption, 2)

If hora < 1 Then
    importe = 200 + (10 * carga) / 100
    lblImp.Caption = importe
End If

If hora > 1 Then
    Select Case min
        Case Is > 45
            importe = hora * (200 + (10 * carga) / 100) + 200 + (10 * carga) / 100
        Case Is > 30
            importe = hora * (200 + (10 * carga) / 100) + 75 * (200 + (10 * carga) / 100) / 100
        Case Is > 15
            importe = hora * (200 + (10 * carga) / 100) + 0.5 * (200 + (10 * carga) / 100)
        Case Is >= 0
            importe = hora * (200 + (10 * carga) / 100) + 0.25 * (200 + (10 * carga) / 100)
    End Select
    lblImp.Caption = importe
End If
    

                     
End Sub

Private Sub Form_Activate()

Text1.Text = Date
Text2.Text = Time

End Sub

