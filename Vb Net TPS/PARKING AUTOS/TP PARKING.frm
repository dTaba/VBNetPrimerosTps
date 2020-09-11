VERSION 5.00
Begin VB.Form TUTÚ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TU-TÚ PARKING"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TP PARKING.frx":0000
   ScaleHeight     =   4245
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRecuperar 
      Caption         =   "Restaurar"
      Height          =   495
      Left            =   4320
      TabIndex        =   10
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CheckBox checkestadia 
      Caption         =   "Estadía"
      Height          =   255
      Left            =   6480
      TabIndex        =   9
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Timer Timer 
      Interval        =   1
      Left            =   480
      Top             =   3480
   End
   Begin VB.TextBox txtCantidad 
      Enabled         =   0   'False
      Height          =   375
      Left            =   7560
      TabIndex        =   6
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton cmdSalida 
      Caption         =   "Salida"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdIngreso 
      Caption         =   "Ingreso"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.ListBox lstIngreso 
      Enabled         =   0   'False
      Height          =   2595
      Left            =   2400
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.ListBox lstSalida 
      Enabled         =   0   'False
      Height          =   2595
      Left            =   4320
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.ListBox lstPatente 
      Enabled         =   0   'False
      Height          =   2595
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   16200
      Left            =   -8040
      Picture         =   "TP PARKING.frx":1DAC7
      Stretch         =   -1  'True
      Top             =   -6480
      Width           =   28800
   End
   Begin VB.Label lblHora 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label lblFecha 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblVehiculos 
      Caption         =   "Vehículos"
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   2640
      Width           =   855
   End
End
Attribute VB_Name = "TUTÚ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cantidad As Integer
Dim caca As Boolean
Dim dato As String
Private Function guardar()

Kill (App.Path & "\backup.txt")
Open (App.Path & "\backup.txt") For Append As #1

For i = 0 To lstPatente.ListCount - 1
    Write #1, lstPatente.List(i); lstIngreso.List(i); lstSalida.List(i)
Next

Close #1

End Function

Private Sub cmdIngreso_Click()

mono = False
caca = False
lol = True
entrante = InputBox("Ingrese la patente del auto entrante.")

If Len(entrante) = 6 Or Len(entrante) = 7 Then
    If Len(entrante) = 7 Then
        If IsNumeric(Trim(UCase(Left(entrante, 2)))) = False And IsNumeric(Trim(UCase(Right(entrante, 2)))) = False And IsNumeric(Trim(UCase(Mid(entrante, 3, 3)))) = True Then
            If lstPatente.ListCount >= 1 Then
                For i = 0 To lstPatente.ListCount - 1
                    If UCase(lstPatente.List(i)) = UCase(entrante) Then
                        lol = False
                        MsgBox "Patente Repetida"
                    End If
                Next
            End If
            If lol = True Then
                lstPatente.AddItem entrante
                lstIngreso.AddItem lblHora.Caption
                If checkestadia.Value = 1 Then
                    lstSalida.AddItem "Estadia"
                Else
                    lstSalida.AddItem "Estacionado"
                End If
                    cantidad = cantidad + 1
                    txtCantidad = cantidad
                    guardar
            End If
        Else
            MsgBox "Ingrese una patente valida"
        End If
    End If
    If Len(entrante) = 6 Then
        For a = 1 To 3
            letra = Trim(UCase(Mid(entrante, a, 1)))
            If letra >= "A" And letra <= "Z" And IsNumeric(Trim(UCase(Right(entrante, 3)))) = True Then
                caca = True
            Else
                caca = False
            End If
        Next
        If caca = True Then
            If lstPatente.ListCount >= 1 Then
                For i = 0 To lstPatente.ListCount - 1
                    If UCase(lstPatente.List(i)) = UCase(entrante) Then
                        MsgBox "Patente Repetida"
                        mono = True
                    End If
                Next
            End If
            If mono = False Then
                lstPatente.AddItem entrante
                lstIngreso.AddItem lblHora.Caption
                If checkestadia.Value = 1 Then
                    lstSalida.AddItem "Estadia"
                Else
                    lstSalida.AddItem "Estacionado"
                End If
                cantidad = cantidad + 1
                txtCantidad = cantidad
                guardar
            End If
        Else
            MsgBox "Ingrese una patente valida"
        End If
    End If
Else
    MsgBox "Ingrese una patente valida."
End If

End Sub

Private Sub cmdRecuperar_Click()

lstPatente.Clear
lstSalida.Clear
lstIngreso.Clear

Open (App.Path & "\backup.txt") For Input As #1
                    
While Not EOF(1)
    Input #1, a, b, c
    lstPatente.AddItem (a)
    lstIngreso.AddItem (b)
    lstSalida.AddItem (c)
    cantidad = 0
    For s = 0 To lstSalida.ListCount - 1
        If lstSalida.List(s) = "Estacionado" Or lstSalida.List(s) = "Estadia" Then
            cantidad = cantidad + 1
        End If
    Next
    txtCantidad = cantidad
Wend

Close #1

End Sub

Private Sub cmdSalida_Click()

pepe = False
saliente = InputBox("Ingrese la patente del auto saliente.")

If Len(saliente) = 6 Or Len(saliente) = 7 Then
    If Len(saliente) = 7 Then
        If IsNumeric(Trim(UCase(Left(saliente, 2)))) = False And IsNumeric(Trim(UCase(Right(saliente, 2)))) = False And IsNumeric(Trim(UCase(Mid(saliente, 3, 3)))) = True Then
            For i = 0 To lstPatente.ListCount - 1
                If UCase(lstPatente.List(i)) = UCase(saliente) And lstSalida.List(i) = "Estacionado" Then
                    lstSalida.List(i) = lblHora.Caption
                    pepe = True
                    intervalo = Format(TimeValue(lstSalida.List(i)) - TimeValue(lstIngreso.List(i)), "hh:mm:ss")
                    frmticket.lblpatente = UCase(saliente)
                    frmticket.lblhinicio = lstIngreso.List(i)
                    frmticket.lblhfinalizacion = lstSalida.List(i)
                    guardar
                End If
                If UCase(lstPatente.List(i)) = UCase(saliente) And lstSalida.List(i) = "Estadia" Then
                    lstSalida.List(i) = lblHora.Caption
                    pepe = True
                    estadia = True
                    intervalo = "Estadia"
                    frmticket.lblpatente = UCase(saliente)
                    frmticket.lblhinicio = lstIngreso.List(i)
                    frmticket.lblhfinalizacion = lstSalida.List(i)
                End If
            Next
        Else
            MsgBox "Vehiculo no encontrado"
        End If
    End If
    If Len(saliente) = 6 Then
        For a = 1 To 3
            letra = Trim(UCase(Mid(saliente, a, 1)))
            If letra >= Chr(65) And letra <= Chr(90) And IsNumeric(Trim(UCase(Right(saliente, 3)))) = True Then
                For i = 0 To lstPatente.ListCount - 1
                    If UCase(lstPatente.List(i)) = UCase(saliente) And lstSalida.List(i) = "Estacionado" Then
                        lstSalida.List(i) = lblHora.Caption
                        pepe = True
                        intervalo = Format(TimeValue(lstSalida.List(i)) - TimeValue(lstIngreso.List(i)), "hh:mm:ss")
                        frmticket.lblpatente = UCase(saliente)
                        frmticket.lblhinicio = lstIngreso.List(i)
                        frmticket.lblhfinalizacion = lstSalida.List(i)
                        guardar
                    End If
                    If UCase(lstPatente.List(i)) = UCase(saliente) And lstSalida.List(i) = "Estadia" Then
                        lstSalida.List(i) = lblHora.Caption
                        pepe = True
                        estadia = True
                        intervalo = "Estadia"
                        frmticket.lblpatente = UCase(saliente)
                        frmticket.lblhinicio = lstIngreso.List(i)
                        frmticket.lblhfinalizacion = lstSalida.List(i)
                    End If
                Next
            End If
        Next
    End If
    If pepe = True Then
        mins = Mid(intervalo, 4, 1)
        horas = Left(intervalo, 2)
        cantidad = cantidad - 1
        txtCantidad = cantidad
        If estadia = True Then
            precio = 130
            frmticket.lblimporte.Text = precio
            frmticket.lblduracion.Text = intervalo
            frmticket.Show
        Else
            If horas = 1 Or horas = 0 Then
                precio = 40
            End If
            If horas = 2 Then
                precio = 80
            End If
            If horas = 3 Then
                precio = 120
            End If
            If mins > 0 And mins < 15 Then
                precio = precio + 10
            End If
            If mins > 15 And mins < 30 Then
                precio = precio + 20
            End If
            If mins > 30 And mins < 45 Then
                precio = precio + 30
            End If
            If mins > 30 And mins < 45 Then
                precio = precio + 40
            End If
            frmticket.lblimporte.Text = "$ " & precio
            frmticket.lblduracion.Text = intervalo
            frmticket.Show
            guardar
        End If
    End If
    If pepe = False Then
        MsgBox "Vehiculo no encontrado"
    End If
Else
    MsgBox "Vehiculo no encontrado"
End If

End Sub

Private Sub Form_Load()

lblFecha.Caption = Date

End Sub

Private Sub Timer_Timer()

lblHora.Caption = Time

End Sub
