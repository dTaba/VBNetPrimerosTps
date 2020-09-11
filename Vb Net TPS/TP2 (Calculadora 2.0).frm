VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3780
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrmBase 
      Caption         =   "Base"
      Height          =   1215
      Left            =   240
      TabIndex        =   15
      Top             =   2400
      Width           =   1695
      Begin VB.OptionButton OptDecimal 
         Caption         =   "Decimal"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton OptBinario 
         Caption         =   "Binario"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Frame FrmResultado 
      Caption         =   "Resultado"
      Height          =   1095
      Left            =   2400
      TabIndex        =   12
      Top             =   2400
      Width           =   2175
      Begin VB.TextBox TxtResultado 
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame FrmDatos 
      Caption         =   "Datos"
      Height          =   1695
      Left            =   2400
      TabIndex        =   7
      Top             =   360
      Width           =   2175
      Begin VB.TextBox TxtDato2 
         Height          =   495
         Left            =   960
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox TxtDato1 
         Height          =   495
         Left            =   960
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Dato 2"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Dato 1"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "Borrar"
      Height          =   615
      Left            =   4920
      TabIndex        =   6
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton CmdCalcular 
      Caption         =   "Calcular"
      Height          =   615
      Left            =   4920
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin VB.Frame FrmOperación 
      Caption         =   "Operación"
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1695
      Begin VB.OptionButton OptDivisión 
         Caption         =   "Cociente"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton OptMultiplicación 
         Caption         =   "Producto"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton OptResta 
         Caption         =   "Resta"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton OptSuma 
         Caption         =   "Suma"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
End Sub

Private Sub CmdBorrar_Click()
    TxtResultado = ""
    TxtDato1 = ""
    TxtDato2 = ""
    OptSuma = False
    OptResta = False
    OptMultiplicación = False
    OptDivisión = False
End Sub

Private Sub CmdCalcular_Click()
    If TxtDato1.Text = "" Or TxtDato2.Text = "" Then
        MsgBox "Faltan datos"
        Exit Sub
    End If
    If IsNumeric(TxtDato1.Text) = False Or IsNumeric(TxtDato2.Text) = False Then
        MsgBox "Datos no válidos"
        Exit Sub
    End If
    If OptDivisión.Value = True Then
        If TxtDato2 = 0 Then
            MsgBox "No es posible dividir por 0"
            Exit Sub
        End If
            TxtResultado.Text = TxtDato1.Text / TxtDato2.Text
    End If
    If OptSuma.Value = True Then
        TxtResultado.Text = Val(TxtDato1.Text) + Val(TxtDato2.Text)
    End If
    If OptResta.Value = True Then
        TxtResultado.Text = TxtDato1.Text - TxtDato2.Text
    End If
    If OptMultiplicación.Value = True Then
        TxtResultado.Text = TxtDato1.Text * TxtDato2.Text
    End If
End Sub

Private Sub CmdSalir_Click()
    rest = MsgBox("¿Desea salir?", vbQuestion + vbYesNo)
    If rest = vbYes Then
        End
    End If
End Sub

Private Sub Form_Load()
    OptSuma = False
    OptResta = False
    OptMultiplicación = False
    OptDivisión = False
End Sub

Private Sub Option1_Click()

End Sub

Private Sub Option2_Click()

End Sub

Private Sub OptBinario_Click()
    If OptBinario.Value = True Then
        Dim Residuo As String 'Declaramos el Residuo
        Dim Numero As Long 'Declaramos la Variable que manejará el numero
        Numero = Val(TxtResultado.Text) ' le damos el valor del textbox a la variable numerica
        TxtResultado.Text = "" 'Seteamos el Resultado a vacio
        Do
            Residuo = Numero Mod 2 'Obtenemos el Residuo de la division
            TxtResultado.Text = TxtResultado.Text & Str(Residuo) 'concatenamos el Residuo al final con lo acumulado en el resultado, recuerden que las cadenas se concatenan con el ampersan "&" no con el "+"
            Numero = Int(Numero / 2) 'Obtenemos el entero de la division
        Loop Until Numero < 2 'Seguimos haciendo la operación hasta que el numero sea 0 o 1
        If (Numero = 1) Then 'verificamos que valor tenemos como ultimo residuo o mejor dicho como ultimo numero
            TxtResultado.Text = "1 " & StrReverse(TxtResultado.Text) 'le agregamos el ultimo valor al inicio ya que el valor anterior lo vamos a revertir
        Else
            TxtResultado.Text = StrReverse(TxtResultado.Text) 'como no hay nada que concatenar, simplemente revertimos
        End If
    End If
End Sub

Private Sub OptDecimal_Click()
If OptDecimal = True Then
    Dim Numero As String 'Declaramos la Variable que manejará el numero como cadena
    Numero = TxtResultado.Text 'le damos el valor del textbox a la variable string
    TxtResultado.Text = "" 'Seteamos el Resultado a vacio
    Dim Total As Long 'Acumulador
    Dim Constante, Temp As Integer 'Constante que irá cambiando en x 2
    Constante = 1 'Iniciamos la contantes en 1
    Do
        Temp = Val(Right(Numero, 1)) 'obtengo el primer numero de la derecha
        Numero = Left(Numero, Len(Numero) - 1) ' Al binario le quito el ultimo digito
        Total = Total + (Temp * Constante) 'El primer digito que saqué de la derecho lo multiplico con la constante
        Constante = Constante * 2 ' la constante será 1,2,4,8,16,32, etc es decir, x 2
    Loop Until Len(Numero) = 0 'Seguimos haciendo la operación hasta la cadena binario se quede sin digitos
    TxtResultado.Text = Total
End If
End Sub
