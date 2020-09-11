VERSION 5.00
Begin VB.Form frmGame 
   Caption         =   "NUMBER GAME"
   ClientHeight    =   3855
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FNivel 
      Caption         =   "Nivel"
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2655
      Begin VB.CommandButton cmdReset 
         Caption         =   "Resetear"
         Height          =   495
         Left            =   1680
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optDesarrollador 
         Caption         =   "Desarrollador"
         CausesValidation=   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   2040
         Width           =   1695
      End
      Begin VB.OptionButton OptDificil 
         Caption         =   "Difícil (1 a 1000)"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   2055
      End
      Begin VB.OptionButton OptIntermedio 
         Caption         =   "Intermedio (1 a 100)"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   1935
      End
      Begin VB.OptionButton OptFacil 
         Caption         =   "Fácil (1 a 10)"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.CommandButton CmdJugar 
      Caption         =   "Jugar"
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox X 
      Height          =   975
      Left            =   3000
      TabIndex        =   9
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox contador 
      Height          =   1095
      Left            =   3000
      TabIndex        =   10
      Top             =   2160
      Width           =   495
   End
   Begin VB.Frame FNumero 
      Caption         =   "Su Número"
      Height          =   2895
      Left            =   3600
      TabIndex        =   5
      Top             =   480
      Width           =   3495
      Begin VB.CommandButton CmdAbandonar 
         Caption         =   "Abandonar"
         Height          =   615
         Left            =   360
         TabIndex        =   8
         Top             =   1920
         Width           =   2775
      End
      Begin VB.CommandButton CmdArriesgar 
         Caption         =   "Arriesgar"
         Height          =   615
         Left            =   360
         TabIndex        =   7
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox TxtNumero 
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Image imgJuego 
      Height          =   5640
      Left            =   0
      Picture         =   "frmGame.frx":0000
      Top             =   0
      Width           =   8550
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdAbandonar_Click()
frmFracasado.Show
End Sub

Private Sub CmdArriesgar_Click()
If TxtNumero.Text = "" Then
 MsgBox ("Debe ingresar un número.")
End If
If IsNumeric(TxtNumero.Text) Then
    If OptFacil.Value = True Then
        If TxtNumero.Text > 10 Or TxtNumero.Text < 1 Then
            MsgBox ("El número se encuentra fuera del rango.")
        Else
        If Val(TxtNumero.Text) < X Then
            MsgBox ("El número es mayor.")
            End If
        If Val(TxtNumero.Text) > X Then
            MsgBox ("El número es menor.")
            End If
            contador2 = contador2 + 1
    End If
    End If
    If OptIntermedio.Value = True Or optDesarrollador.Value = True Then
        If TxtNumero.Text > 100 Or TxtNumero.Text < 1 Then
            MsgBox ("El número se encuentra fuera del rango.")
        Else
        If Val(TxtNumero.Text) < X Then
            MsgBox ("El número es mayor.")
            
        End If
        If Val(TxtNumero.Text) > X Then
            MsgBox ("El número es menor.")
            
        End If
        contador2 = contador2 + 1
        contador.Text = contador2
        
    End If
    End If
    If OptDificil.Value = True Then
        If TxtNumero.Text > 1000 Or TxtNumero.Text < 1 Then
            MsgBox ("El número se encuentra fuera del rango.")
        Else
        If Val(TxtNumero.Text) < X Then
            MsgBox ("El número es mayor.")
            
        End If
        If Val(TxtNumero.Text) > X Then
            MsgBox ("El número es menor.")
            
        End If
        contador2 = contador2 + 1
    End If
    End If
    

    
    
    
    If TxtNumero.Text = X And contador2 = 1 Then
        TxtNumero.Text = ""
        CmdJugar.Enabled = True
        CmdArriesgar.Enabled = False
        CmdAbandonar.Enabled = False
        OptFacil.Enabled = True
        OptIntermedio.Enabled = True
        OptDificil.Enabled = True
        OptFacil.Value = False
        OptIntermedio.Value = False
        OptDificil.Value = False
        optDesarrollador.Value = False
        optDesarrollador.Enabled = True
        X.Visible = False
        contador.Visible = False
        frmWinner.Show
        frmGame.Hide
    End If
    If TxtNumero.Text = X And contador2 > 1 Then
        TxtNumero.Text = ""
        CmdJugar.Enabled = True
        CmdArriesgar.Enabled = False
        CmdAbandonar.Enabled = False
        OptFacil.Enabled = True
        OptIntermedio.Enabled = True
        OptDificil.Enabled = True
        OptFacil.Value = False
        OptIntermedio.Value = False
        OptDificil.Value = False
        optDesarrollador.Value = False
        optDesarrollador.Enabled = True
        X.Visible = False
        contador.Visible = False
        frmGame.Hide
        frmGanador.Show
    End If

End If

End Sub

Private Sub CmdJugar_Click()
contador.Text = 0
If OptFacil.Value = True Then
    Max = 10
    Randomize Timer
    X = Int(Rnd(1) * Max) + Min
End If
If OptIntermedio.Value = True Then
    Max = 100
    Randomize Timer
    X = Int(Rnd(1) * Max) + Min
End If
If OptDificil.Value = True Then
    Max = 1000
    Randomize Timer
    X = Int(Rnd(1) * Max) + Min
End If
If optDesarrollador.Value = True Then

    Max = 100
    Randomize Timer
    X = Int(Rnd(1) * Max) + Min
End If
    If OptFacil.Value = False And OptIntermedio.Value = False And OptDificil.Value = False And optDesarrollador.Value = False Then
        MsgBox "Seleccione una dificultad"
Else
        contador2 = 0
        nombre = InputBox("¿Cuál es su nombre?")
End If


    If Not nombre = "" Then
        contador.Visible = True
    X.Visible = True
    CmdJugar.Enabled = False
        CmdArriesgar.Enabled = True
        CmdAbandonar.Enabled = True
        OptFacil.Enabled = False
        OptIntermedio.Enabled = False
        OptDificil.Enabled = False
        optDesarrollador.Enabled = False
Else
        MsgBox "Ingrese un Nombre"
End If










End Sub

Private Sub cmdReset_Click()
MsgBox "Desea reiniciar todos los datos?", vbYesNo, "Resetear"
If vbYes Then
        TxtNumero.Text = ""
        CmdJugar.Enabled = True
        CmdArriesgar.Enabled = False
        CmdAbandonar.Enabled = False
        OptFacil.Enabled = True
        OptIntermedio.Enabled = True
        OptDificil.Enabled = True
        optDesarrollador.Enabled = True
        optDesarrollador.Value = False
        OptFacil.Value = False
        OptIntermedio.Value = False
        OptDificil.Value = False
        nombre = ""
        contador.Visible = False
        X.Visible = False
        contador2 = 0
        End If
End Sub

Private Sub Form_Activate()
optDesarrollador.Value = False
End Sub

Private Sub Form_Load()
contador = 0
Min = 1
OptFacil.Value = False
OptIntermedio.Value = False
OptDificil.Value = False
optDesarrollador = False

CmdArriesgar.Enabled = False
CmdAbandonar.Enabled = False
contador.Visible = False
X.Visible = False

End Sub

Private Sub OptDificil_Click()
OptFacil.Value = False
OptIntermedio.Value = False
OptDificil.Value = True
End Sub

Private Sub OptFacil_Click()
OptFacil.Value = True
OptIntermedio.Value = False
OptDificil.Value = False
End Sub

Private Sub OptIntermedio_Click()
OptFacil.Value = False
OptIntermedio.Value = True
OptDificil.Value = False
End Sub

Private Sub Option1_Click()

End Sub
