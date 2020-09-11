VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   5220
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Numero7 
      Caption         =   "7"
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton Numero9 
      Caption         =   "9"
      Height          =   495
      Left            =   1560
      TabIndex        =   21
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton Multiplicar 
      Caption         =   "*"
      Height          =   495
      Left            =   2280
      TabIndex        =   20
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton Dividir 
      Caption         =   "/"
      Height          =   495
      Left            =   2280
      TabIndex        =   19
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton Numero8 
      Caption         =   "8"
      Height          =   495
      Left            =   840
      TabIndex        =   18
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton Numero4 
      Caption         =   "4"
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton Borrar 
      Caption         =   "<-"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3000
      TabIndex        =   16
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Numero5 
      Caption         =   "5"
      Height          =   495
      Left            =   840
      TabIndex        =   15
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton Numero6 
      Caption         =   "6"
      Height          =   495
      Left            =   1560
      TabIndex        =   14
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton Numero0 
      Caption         =   "0"
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton Menos 
      Caption         =   "-"
      Height          =   495
      Left            =   2280
      TabIndex        =   12
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton Igual 
      Caption         =   "="
      Height          =   495
      Left            =   840
      TabIndex        =   11
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton C 
      Caption         =   "C"
      Height          =   1335
      Left            =   3000
      TabIndex        =   10
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton Mas 
      Caption         =   "+"
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Numero3 
      Caption         =   "3"
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Numero2 
      Caption         =   "2"
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Numero1 
      Caption         =   "1"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton MMENOS 
      Caption         =   "M-"
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton MMAS 
      Caption         =   "M+"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton MS 
      Caption         =   "MS"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton MR 
      Caption         =   "MR"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton MC 
      Caption         =   "MC"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox Display 
      Alignment       =   1  'Right Justify
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dato1 As Long
Dim dato2 As Long
Dim resultado As String
Dim Dividiractivo As Boolean

Private Sub Dividir_Click()
dato1 = Display.Text
Dividiractivo = True
Display.Text = ""
End Sub

Private Sub Igual_Click()

If Dividiractivo = True Then
resultado = dato1 / Display.Text

End If
Display.Text = resultado
End Sub
Private Sub Numero0_Click()
Display.Text = Display.Text + "0"
If Dividiractivo = True Then
MsgBox "No se puede dividir por 0"
Display.Text = ""
End If
End Sub
Private Sub Numero1_Click()
Display.Text = Display.Text + "1"
End Sub
Private Sub Numero2_Click()
Display.Text = Display.Text + "2"
End Sub
Private Sub Numero3_Click()
Display.Text = Display.Text + "3"
End Sub
Private Sub Numero4_Click()
Display.Text = Display.Text + "4"
End Sub
Private Sub Numero5_Click()
Display.Text = Display.Text + "5"
End Sub
Private Sub Numero6_Click()
Display.Text = Display.Text + "6"
End Sub
Private Sub Numero7_Click()
Display.Text = Display.Text + "7"
End Sub
Private Sub Numero8_Click()
Display.Text = Display.Text + "8"
End Sub
Private Sub Numero9_Click()
Display.Text = Display.Text + "9"
End Sub
