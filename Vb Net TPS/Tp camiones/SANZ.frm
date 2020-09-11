VERSION 5.00
Begin VB.Form frmPrincipal 
   Caption         =   "CACHO S.R.L"
   ClientHeight    =   5370
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12555
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   12555
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstRegreso 
      Height          =   840
      Left            =   9600
      TabIndex        =   18
      Top             =   4080
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ListBox lstCarga 
      Height          =   840
      Left            =   9600
      TabIndex        =   17
      Top             =   2880
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ListBox lstSalida 
      Height          =   840
      Left            =   9600
      TabIndex        =   16
      Top             =   1800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ListBox lstDom2 
      Height          =   840
      Left            =   9600
      TabIndex        =   15
      Top             =   720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   240
      Top             =   3240
   End
   Begin VB.CommandButton cmdBaja 
      Caption         =   "Baja"
      Height          =   615
      Left            =   7200
      TabIndex        =   11
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmdAlta 
      Caption         =   "Alta"
      Height          =   615
      Left            =   7200
      TabIndex        =   10
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox txtDominio 
      Enabled         =   0   'False
      Height          =   495
      Left            =   4200
      TabIndex        =   9
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox txtCarga 
      Height          =   495
      Left            =   4200
      TabIndex        =   8
      Top             =   3600
      Width           =   2175
   End
   Begin VB.ListBox lstDisp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   4800
      TabIndex        =   5
      Top             =   600
      Width           =   1815
   End
   Begin VB.ListBox lstCMax 
      Height          =   2595
      Left            =   2520
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
   Begin VB.ListBox lstDom 
      Height          =   2595
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "Regresar"
      Height          =   615
      Left            =   7080
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Limpiar datos"
      Height          =   615
      Left            =   7080
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   615
      Left            =   7080
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Left            =   3960
      TabIndex        =   14
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblFecha 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblFecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   13
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label lblHora 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblHora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Dominio Seleccionado"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Carga A Transportar"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   3720
      Width           =   1575
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim caca As Boolean

Private Sub cmdAlta_Click()

nuevodom = InputBox("Ingrese el nuevo dominio")


If IsNumeric(Trim(UCase(Left(nuevodom, 3)))) = False And IsNumeric(Trim(UCase(Right(nuevodom, 3)))) = True And Len(nuevodom) = 6 Then
    nuevacap = InputBox("Ingrese la capacidad del nuevo camión")
    If IsNumeric(nuevacap) = True Then
        lstCMax.AddItem nuevacap
        lstDom.AddItem nuevodom
        lstDisp.AddItem "Si"
        For i = 0 To lstCMax.ListCount - 2
            For j = i + 1 To lstCMax.ListCount - 1
                If Val(lstCMax.List(i)) > Val(lstCMax.List(j)) Then
                    aux = lstCMax.List(i)
                    lstCMax.List(i) = lstCMax.List(j)
                    lstCMax.List(j) = aux
                    aux = lstDom.List(i)
                    lstDom.List(i) = lstDom.List(j)
                    lstDom.List(j) = aux
                    aux = lstDisp.List(i)
                    lstDisp.List(i) = lstDisp.List(j)
                    lstDisp.List(j) = aux
                End If
            Next
        Next
    Else: MsgBox "DATOS INVALIDOS"
    End If
Else: MsgBox "DATOS INVALIDOS"
End If

End Sub

Private Sub cmdBaja_Click()

rip = MsgBox("Desea dar de baja el camion seleccionado?", vbQuestion + vbYesNo)
banana = lstDisp.ListIndex

If rip = vbYes Then
    If Not lstDisp.ListIndex = -1 Then
        If lstDisp.List(banana) = "Si" Then
            lstDisp.RemoveItem (lstDom.ListIndex)
            lstCMax.RemoveItem (lstDom.ListIndex)
            lstDom.RemoveItem (lstDom.ListIndex)
        Else
            MsgBox "El camión está siendo usado"
        End If
    End If
End If

End Sub

Private Sub cmdBuscar_Click()

If Not txtCarga.Text = "" Then
    If IsNumeric(txtCarga.Text) = True Then
        If txtCarga.Text > 0 Then
            For a = 0 To Val(lstCMax.ListCount - 1)
                If caca = True Then
                    If lstDisp.List(a) = "Si" Then
                        If Val(txtCarga.Text) <= Val(lstCMax.List(a)) Then
                            txtDominio.Text = lstDom.List(a)
                            caca = False
                            lstDom.ListIndex = a
                            lstCMax.ListIndex = a
                            lstDisp.ListIndex = a
                            lele = MsgBox("Desea enviar este camion?", vbQuestion + vbYesNo)
                            If lele = vbYes Then
                                lstDisp.List(lstCMax.ListIndex) = "No"
                                salida = Time
                                dominio = txtDominio.Text
                                carga = txtCarga.Text
                            End If
                        End If
                    End If
                End If
            Next
        Else
            MsgBox "Ingrese un número positivo"
        End If
    Else
        MsgBox "Ingrese únicamente valores numéricos"
    End If
Else
    MsgBox "Ingrese datos"
End If

Label3.Caption = salida
lstDom2.AddItem (txtDominio.Text)
lstSalida.AddItem (salida)
lstCarga.AddItem (txtCarga.Text)

End Sub

Private Sub cmdLimpiar_Click()

txtCarga.Text = ""
txtDominio.Text = ""

End Sub

Private Sub cmdRegresar_Click()

domreg = InputBox("Ingrese el dominio del camión a regresar")

If IsNumeric(Right(domreg, 3)) = True And IsNumeric(Left(domreg, 3)) = False And Len(domreg) = 6 Then
    For i = 0 To lstDom.ListCount - 1
        If lstDom.List(i) = UCase(domreg) Then
            If lstDisp.List(i) = "No" Then
                lstDisp.List(i) = "Si"
                lstDom.ListIndex = i
                lstCMax.ListIndex = i
                lstDisp.ListIndex = i
            Else
                MsgBox "El camión ya se encuentra disponible"
            End If
        End If
    Next
Else
    MsgBox "Ingrese un dominio válido"
End If

For X = 0 To lstDom.ListCount - 1
    If UCase(domreg) = lstDom2.List(X) Then
        regreso = Time
        frmTicket.lblDom.Caption = lstDom2.List(X)
        frmTicket.lblCarga.Caption = lstCarga.List(X)
        frmTicket.txtHI.Text = lstSalida.List(X)
        frmTicket.txtHF.Text = regreso
        tiempo = Format(TimeValue(frmPrincipal.lstSalida.List(X)) - TimeValue(regreso), "hh:mm")
        frmTicket.lblTiempo.Caption = tiempo
        hora = Left(lblHora.Caption, 2)
        min = Right(lblHora.Caption, 2)
        If hora < 1 Then
            importe = 200 + (10 * carga) / 100
            frmTicket.lblImp.Caption = importe
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
            frmTicket.lblImp.Caption = importe
        End If
        frmTicket.Show
        lstDom2.RemoveItem (X)
        lstCarga.RemoveItem (X)
        lstSalida.RemoveItem (X)
    End If
Next

hora = 0
min = 0

End Sub



Private Sub lstCMax_Click()

lstDom.ListIndex = lstCMax.ListIndex
lstDisp.ListIndex = lstCMax.ListIndex

End Sub

Private Sub lstDisp_Click()

lstDom.ListIndex = lstDisp.ListIndex
lstCMax.ListIndex = lstDisp.ListIndex

End Sub

Private Sub lstDom_Click()

lstCMax.ListIndex = lstDom.ListIndex
lstDisp.ListIndex = lstDom.ListIndex

End Sub

Private Sub Form_Load()
 
lblFecha.Caption = Date

lstDom.AddItem "KEJ151"
lstDom.AddItem "PRS823"
lstDom.AddItem "NAZ421"
lstDom.AddItem "JCT576"
lstDom.AddItem "MAC493"

lstCMax.AddItem "3000"
lstCMax.AddItem "5000"
lstCMax.AddItem "8000"
lstCMax.AddItem "10000"
lstCMax.AddItem "12000"

lstDisp.AddItem "Si"
lstDisp.AddItem "Si"
lstDisp.AddItem "Si"
lstDisp.AddItem "Si"
lstDisp.AddItem "Si"

End Sub


Private Sub Timer_Timer()

lblHora.Caption = Time

End Sub

Private Sub txtCarga_Change()

caca = True

End Sub

