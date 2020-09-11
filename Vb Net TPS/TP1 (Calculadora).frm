VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4635
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   12045
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraMemoria 
      Caption         =   "Memoria"
      Height          =   855
      Left            =   7440
      TabIndex        =   13
      Top             =   1800
      Width           =   4095
      Begin VB.CommandButton CmdMC 
         Caption         =   "MC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdMR 
         Caption         =   "MR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdGuardar 
         Caption         =   "M+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "1 / x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10200
      TabIndex        =   12
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton CmdRaíz 
      Caption         =   "Raíz"
      Height          =   615
      Left            =   9120
      TabIndex        =   11
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton CmdCuadrado 
      Caption         =   " x² "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   10
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   10680
      TabIndex        =   9
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "Borrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   840
      TabIndex        =   8
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton CmdDividir 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      TabIndex        =   6
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton CmdMultiplicar 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   5
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton CmdRestar 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      TabIndex        =   4
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton CmdSumar 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   3
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox TxtDato2 
      Height          =   855
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox TxtDato1 
      Height          =   855
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resultado"
      Height          =   1335
      Left            =   3600
      TabIndex        =   0
      Top             =   720
      Width           =   3495
      Begin VB.TextBox TxtResultado 
         Height          =   975
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   2160
      Top             =   720
      Width           =   15
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mem As Single

Private Sub Command1_Click()

End Sub

Private Sub Cmd1_Click()
If TxtDato1.Text = "" Then
MsgBox "Faltan datos"
Exit Sub
End If
If IsNumeric(TxtDato1.Text) = False Then
MsgBox "Datos no válidos"
Exit Sub
End If
TxtResultado = 1 / TxtDato1
End Sub

Private Sub CmdBorrar_Click()
TxtResultado = ""
TxtDato1 = ""
TxtDato2 = ""
End Sub

Private Sub CmdCuadrado_Click()
If TxtDato1.Text = "" Then
MsgBox "Faltan datos"
Exit Sub
End If
If IsNumeric(TxtDato1.Text) = False Then
MsgBox "Datos no válidos"
Exit Sub
End If
TxtResultado = TxtDato1 ^ 2
End Sub

Private Sub CmdDividir_Click()
If TxtDato1.Text = "" Or TxtDato2.Text = "" Then
MsgBox "Faltan datos"
Exit Sub
End If
If IsNumeric(TxtDato1.Text) = False Or IsNumeric(TxtDato2.Text) = False Then
MsgBox "Datos no válidos"
Exit Sub
End If
If TxtDato2 = 0 Then
    MsgBox "No es posible dividir por 0"
    Exit Sub
End If
TxtResultado.Text = TxtDato1.Text / TxtDato2.Text
End Sub


Private Sub CmdMC_Click()
Mem = 0
CmdGuardar.BackColor = &H8000000F
End Sub

Private Sub CmdGuardar_Click()
If TxtResultado.Text = "" Then
MsgBox "Faltan datos"
Exit Sub
End If
Mem = TxtResultado.Text
If Mem <> 0 Then
    CmdGuardar.BackColor = vbBlue
End If
End Sub

Private Sub CmdMR_Click()
TxtDato1 = Mem
End Sub

Private Sub CmdMultiplicar_Click()
If TxtDato1.Text = "" Or TxtDato2.Text = "" Then
MsgBox "Faltan datos"
Exit Sub
End If
If IsNumeric(TxtDato1.Text) = False Or IsNumeric(TxtDato2.Text) = False Then
MsgBox "Datos no válidos"
Exit Sub
End If
TxtResultado.Text = TxtDato1.Text * TxtDato2.Text
End Sub

Private Sub CmdRaíz_Click()
If TxtDato1.Text = "" Then
MsgBox "Faltan datos"
Exit Sub
End If
If IsNumeric(TxtDato1.Text) = False Then
MsgBox "Datos no válidos"
Exit Sub
End If
TxtResultado = Sqr(TxtDato1)
End Sub

Private Sub CmdRestar_Click()
If TxtDato1.Text = "" Or TxtDato2.Text = "" Then
MsgBox "Faltan datos"
Exit Sub
End If
If IsNumeric(TxtDato1.Text) = False Or IsNumeric(TxtDato2.Text) = False Then
MsgBox "Datos no válidos"
Exit Sub
End If
TxtResultado.Text = TxtDato1.Text - TxtDato2.Text
End Sub

Private Sub CmdSalir_Click()
End
End Sub

Private Sub CmdSumar_Click()
If TxtDato1.Text = "" Or TxtDato2.Text = "" Then
MsgBox "Faltan datos"
Exit Sub
End If
If IsNumeric(TxtDato1.Text) = False Or IsNumeric(TxtDato2.Text) = False Then
MsgBox "Datos no válidos"
Exit Sub
End If
TxtResultado.Text = Val(TxtDato1.Text) + Val(TxtDato2.Text)
End Sub

