VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmTipoMovInv 
   Caption         =   "Tipos de Movimientos de Inventarios"
   ClientHeight    =   8775
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14970
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   14970
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   960
      TabIndex        =   5
      Top             =   0
      Width           =   11895
      Begin VB.Frame Frame1 
         Height          =   1215
         Left            =   360
         TabIndex        =   14
         Top             =   960
         Width           =   8895
         Begin VB.OptionButton optSalida 
            Caption         =   "Salida"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5760
            TabIndex        =   19
            Top             =   720
            Width           =   1815
         End
         Begin VB.OptionButton optEntrada 
            Caption         =   "Entrada"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2880
            TabIndex        =   18
            Top             =   720
            Width           =   1935
         End
         Begin VB.OptionButton optResta 
            Caption         =   "Negativo (Resta factor=  -1  )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5760
            TabIndex        =   16
            Top             =   240
            Width           =   3015
         End
         Begin VB.OptionButton optSuma 
            Caption         =   "Positivo (Suma factor=1)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2880
            TabIndex        =   15
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo de Movimiento :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   20
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label4 
            Caption         =   "Efecto en Inventario :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   17
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.TextBox txtTransaccion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10680
         TabIndex        =   13
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtIDTipo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtDescr 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3840
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   600
         Width           =   4095
      End
      Begin VB.CheckBox chkReservada 
         Caption         =   "Reservada por el Sistema ( Read Only )"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   9600
         TabIndex        =   6
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Transacción Abreviada :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8280
         TabIndex        =   12
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "IDTipo :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdEliminar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15120
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Elimina el item actualmente seleccionado en el grid de datos ..."
      Top             =   4200
      Width           =   810
   End
   Begin VB.CommandButton cmdAdd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15120
      Picture         =   "Form1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Agrega el item con los datos digitados..."
      Top             =   4920
      Width           =   810
   End
   Begin VB.CommandButton cmdSave 
      Enabled         =   0   'False
      Height          =   495
      Left            =   15120
      Picture         =   "Form1.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Aplica y Guarda los datos de la transacción en Firme ..."
      Top             =   5640
      Width           =   795
   End
   Begin VB.CommandButton cmdEditItem 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15120
      Picture         =   "Form1.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Modifica los datos mostrados en el Grid con los datos digitados ..."
      Top             =   3480
      Width           =   810
   End
   Begin VB.CommandButton cmdUndo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15120
      Picture         =   "Form1.frx":1320
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Deshacer / Cancelar"
      Top             =   6360
      Width           =   810
   End
   Begin TrueOleDBGrid60.TDBGrid TDBG 
      Height          =   5055
      Left            =   480
      OleObjectBlob   =   "Form1.frx":1762
      TabIndex        =   11
      Top             =   2880
      Width           =   14865
   End
End
Attribute VB_Name = "frmTipoMovInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As ADODB.Recordset
Dim bOrdenCodigo As Boolean
Dim bOrdenDescr As Boolean
Dim sCodSucursal As String
Dim bEdit As Boolean
Dim bAdd As Boolean
Dim sSoloActivo As String

Private Sub chkActivo_Click()
If bAdd Then
    chkActivo.value = 1
End If
End Sub

Private Sub chkIndependiente_Click()
If bAdd And chkIndependiente.value = 1 Then
    txtVendedor.Text = ""
End If
If bAdd And chkIndependiente.value = 0 Then
    txtVendedor.Text = "ND"
End If

End Sub

Private Sub chkSoloActivos_Click()
If chkSoloActivos.value = 1 Then
    sSoloActivo = "1"
Else
    sSoloActivo = "0"
End If
End Sub

Private Sub chkAfectaVacaciones_Click()
If chkAfectaVacaciones.value = 1 Then
    optResta.value = False
    optSuma.value = False
End If

End Sub

Private Sub chkCalcVacacional_Click()
'Dim lbok As Boolean
'
'If chkCalcVacacional.value = 1 And (chkSaldoInicialPeriodo.value = 1 Or chkSaldoSalidaEmpleado.value = 1) Then
'    lbok = Mensaje("Se ha indicado cálculo vacacional pero también como Saldo de Período y/o Saldo x Salida Empleado. Son Excluyentes !!!", ICO_ERROR, False)
'    chkCalcVacacional.value = 0
'
'End If

End Sub



Private Sub chkRRHH_Click()
If bAdd Then
    If chkRRHH.value = 0 Then
        txtTipoAccion.Enabled = True
        fmtTextbox txtTipoAccion, "O"
        txtDescr.Enabled = True
        fmtTextbox txtDescr, "O"
    Else
        txtTipoAccion.Enabled = False
        fmtTextbox txtTipoAccion, "R"
        txtDescr.Enabled = False
        fmtTextbox txtDescr, "R"
    
    End If
End If
End Sub

Private Sub chkSaldoInicialPeriodo_Click()
If chkSaldoInicialPeriodo.value = 1 And (chkCalcVacacional.value = 1 Or chkSaldoSalidaEmpleado.value = 1) Then
    lbok = Mensaje("Se ha indicado Saldo de Período  pero también como cálculo vacacional y/o Saldo x Salida Empleado. Son Excluyentes !!!", ICO_ERROR, False)
    chkSaldoInicialPeriodo.value = 0
    
End If
End Sub

Private Sub chkSaldoSalidaEmpleado_Click()
If chkSaldoSalidaEmpleado.value = 1 And (chkCalcVacacional.value = 1 Or chkSaldoInicialPeriodo.value = 1) Then
    lbok = Mensaje("Se ha indicado Saldo x Salida Empleado  pero también como cálculo vacacional y/o Saldo de Período. Son Excluyentes !!!", ICO_ERROR, False)
    chkSaldoSalidaEmpleado.value = 0
    
End If
End Sub

Private Sub cmdAdd_Click()
bAdd = True
bEdit = False
txtPrioridad.Enabled = True
txtPrioridad.Text = ""
cmdTipo.Enabled = True
chkActivo.Enabled = True
chkActivo.value = 1
chkRRHH.Enabled = False
chkRRHH.value = 0
chkSaldoInicialPeriodo.Enabled = True
chkSaldoInicialPeriodo.value = 0
chkSaldoSalidaEmpleado.Enabled = True
chkAfectaVacaciones.Enabled = True
chkAfectaVacaciones.value = 0
optSuma.Enabled = True
optSuma.value = False
optResta.Enabled = True
optResta.value = False
chkCalcVacacional.Enabled = True
chkCalcVacacional.value = 0
chkRepiteMes.Enabled = True
chkRepiteMes.value = 0
chkUnaVezMes.Enabled = True
chkUnaVezMes.value = 0
chkUnaVezPeriodo.Enabled = True
chkUnaVezPeriodo.value = 0
chkAjuste.Enabled = True
chkAjuste.value = 0
chkMaximoUno.Enabled = True
chkMaximoUno.value = 0
fmtTextbox txtTipoAccion, "O"
txtTipoAccion.Text = ""
fmtTextbox txtDescr, "O"
txtDescr.Text = ""
cmdSave.Enabled = True
cmdEliminar.Enabled = False
cmdAdd.Enabled = False
'txtTipoAccion.SetFocus
End Sub

Private Sub cmdEditItem_Click()
bEdit = True
bAdd = False
GetDataFromGridToControl
cmdTipo.Enabled = False
chkActivo.Enabled = True
chkRRHH.Enabled = False
chkAfectaVacaciones.Enabled = True
optSuma.Enabled = True
optResta.Enabled = True
chkCalcVacacional.Enabled = True
chkSaldoInicialPeriodo.Enabled = True
chkSaldoSalidaEmpleado.Enabled = True
chkUnaVezPeriodo.Enabled = True
chkRepiteMes.Enabled = True
chkUnaVezMes.Enabled = True
chkAjuste.Enabled = True
chkMaximoUno.Enabled = True
fmtTextbox txtTipoAccion, "R"
fmtTextbox txtDescr, "O"
txtPrioridad.Enabled = True
cmdSave.Enabled = True
cmdEliminar.Enabled = False
cmdAdd.Enabled = False
End Sub

Private Sub initControles()
txtPrioridad.Enabled = False
cmdTipo.Enabled = False
chkActivo.Enabled = False
chkRRHH.Enabled = False
chkSaldoInicialPeriodo.Enabled = False
chkSaldoSalidaEmpleado.Enabled = False
chkAfectaVacaciones.Enabled = False
chkUnaVezMes.Enabled = False
chkUnaVezPeriodo.Enabled = False
chkRepiteMes.Enabled = False
optSuma.Enabled = False
optResta.Enabled = False
chkCalcVacacional.Enabled = False
chkAjuste.Enabled = False
chkMaximoUno.Enabled = False
fmtTextbox txtTipoAccion, "R"
fmtTextbox txtDescr, "R"
End Sub

Private Sub GetDataFromGridToControl()
If Not (rst.EOF And rst.BOF) Then
    txtTipoAccion.Text = rst("Tipo_Accion").value
    txtDescr.Text = rst("Descr").value
    txtPrioridad.Text = rst("Prioridad").value
    If rst("flgRRHH").value = "SI" Then
        chkRRHH.value = 1
    Else
        chkRRHH.value = 0
    End If
    If rst("AfectaVacaciones").value = "SI" Then
        chkAfectaVacaciones.value = 1
        If rst("Factor").value = "SUMA" Then
            optSuma.value = True
        End If
        If rst("Factor").value = "RESTA" Then
            optResta.value = True
        End If
    Else
    
        chkAfectaVacaciones.value = 0
        optSuma.value = False
        optResta.value = True
    End If
        
    If rst("CalculoVacacional").value = "SI" Then
        chkCalcVacacional.value = 1
    Else
        chkCalcVacacional.value = 0
    End If
    If rst("SaldoSalidaEmpleado").value = "SI" Then
        chkSaldoSalidaEmpleado.value = 1
    Else
        chkSaldoSalidaEmpleado.value = 0
    End If
        
    If rst("SaldoInicialPeriodo").value = "SI" Then
        chkSaldoInicialPeriodo.value = 1
    Else
        chkSaldoInicialPeriodo.value = 0
    End If
        
    If rst("EsAjuste").value = "SI" Then
        chkAjuste.value = 1
    Else
        chkAjuste.value = 0
    End If
        
    If rst("Activo").value = "SI" Then
        chkActivo.value = 1
    Else
        chkActivo.value = 0
    End If
    
    If rst("SeRepiteMes").value = "SI" Then
        chkRepiteMes.value = 1
    Else
        chkRepiteMes.value = 0
    End If
    
    If rst("UnaVezPeriodo").value = "SI" Then
        chkUnaVezPeriodo.value = 1
    Else
        chkUnaVezPeriodo.value = 0
    End If
    
    If rst("UnaVezMes").value = "SI" Then
        chkUnaVezMes.value = 1
    Else
        chkUnaVezMes.value = 0
    End If
    
    If rst("ValorMaximoUno").value = "SI" Then
        chkMaximoUno.value = 1
    Else
        chkMaximoUno.value = 0
    End If
    
End If

End Sub


Private Sub cmdEliminar_Click()
Dim lbok As Boolean
Dim sMsg As String
Dim sTipo As String
Dim sFiltro As String

    If txtTipoAccion.Text = "" Then
        lbok = Mensaje("El Código de la Acción no pueden estar en Blanco", ICO_ERROR, False)
        Exit Sub
    End If
    
    If ExistCodeInTable("TipoSaldo", True, txtTipoAccion.Text, "sgvSaldoDiasEmpleado", False, "") Then
        lbok = Mensaje("La Acción tiene Movimientos, no se puede eliminar", ICO_ADVERTENCIA, False)
        Exit Sub
    End If
    
    If txtDescr.Text = "" Then
        lbok = Mensaje("La Descripción no puede estar en Blanco", ICO_ERROR, False)
        Exit Sub
    End If
    If txtPrioridad.Text = "" Then
        lbok = Mensaje("La prioridad no puede estar en Blanco", ICO_ERROR, False)
        Exit Sub
    End If

    lbok = Mensaje("Está seguro de eliminar esa Acción " & rst("Descr").value, ICO_PREGUNTA, True)
    If lbok Then
                lbok = sgvActualizaTipoAccion(txtTipoAccion.Text, txtDescr.Text, "0", "0", "0", _
        "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "D")
        
        If lbok Then
            sMsg = "Borrado Exitosamente ... "
            lbok = Mensaje(sMsg, ICO_OK, False)
            ' actualiza datos
            cargaGridTiposAccion "2"
        End If
    End If
End Sub


Private Sub cmdSave_Click()

Dim lbok As Boolean
Dim sMsg As String
Dim sTipo As String
Dim sFiltro As String
Dim sTipoDescr As String
Dim sIndependiente As String
Dim sActivo As String
Dim schkRRHH As String
Dim schkAfecta As String
Dim sFactor As String 'optSuma , optResta(Factor)
Dim schkActivo As String
Dim schkRecalculoVac As String
Dim schkAjuste As String
Dim stmpValorSINO As String
Dim sEsSaldoInicialPeriodo As String
Dim sSaldoSalidaEmpleado As String
Dim sSeRepiteMes As String
Dim sUnaVezPeriodo As String
Dim sUnaVezMes As String
Dim sMaximoUno As String


    If txtTipoAccion.Text = "" Then
        lbok = Mensaje("El Código de la Acción no pueden estar en Blanco", ICO_ERROR, False)
        Exit Sub
    End If
    If txtDescr.Text = "" Then
        lbok = Mensaje("La Descripción no puede estar en Blanco", ICO_ERROR, False)
        Exit Sub
    End If
    If txtPrioridad.Text = "" Then
        lbok = Mensaje("La prioridad no puede estar en Blanco", ICO_ERROR, False)
        Exit Sub
    End If

    If Not Val_TextboxNum(txtPrioridad) Then
        lbok = Mensaje("La prioridad debe ser numérica", ICO_ERROR, False)
        Exit Sub
    End If
    
    If (chkUnaVezMes.value + Me.chkUnaVezPeriodo.value + chkRepiteMes.value) = 0 Then
        lbok = Mensaje("La periodicidad de la acción no ha sido especificada, debe indicar si repite mes, repite período ó una vez al mes", ICO_ERROR, False)
        Exit Sub
    End If
    
    
    If (chkUnaVezMes.value + Me.chkUnaVezPeriodo.value + chkRepiteMes.value) > 1 Then
        lbok = Mensaje("La periodicidad de la acción es incorrecta, son excluyente repite mes, repite período y una vez al mes", ICO_ERROR, False)
        Exit Sub
    End If
    
    If chkAfectaVacaciones.value = 1 Then
        schkAfecta = "1"
    Else
        schkAfecta = "0"
    End If
    
    If chkSaldoInicialPeriodo.value = 1 Then
        sEsSaldoInicialPeriodo = "1"
    Else
        sEsSaldoInicialPeriodo = "0"
    End If
    
    
    If chkSaldoSalidaEmpleado = 1 Then
        sSaldoSalidaEmpleado = "1"
    Else
        sSaldoSalidaEmpleado = "0"
    End If
    
    If chkAfectaVacaciones.value = 1 And (optSuma.value = 0 And optResta.value = 0) Then
            lbok = Mensaje("La acción afecta vacaciones y no se ha especificado si suma o resta a las vacaciones", ICO_ERROR, False)
        Exit Sub

    End If
    sFactor = "0"
    If chkAfectaVacaciones.value = 1 And (optSuma.value = True Or optResta.value = True) Then
        If optSuma.value = True Then
            sFactor = "1"
        End If
        If optResta.value = True Then
            sFactor = "-1"
        End If

    End If
    
    If chkCalcVacacional.value = 1 Then
        schkRecalculoVac = "1"
    Else
        schkRecalculoVac = "0"
    End If
    
    If chkAjuste.value = 1 Then
        schkAjuste = "1"
    Else
        schkAjuste = "0"
    End If
    
    If chkRepiteMes.value = 1 Then
        sSeRepiteMes = "1"
    Else
        sSeRepiteMes = "0"
    End If
    
    If chkUnaVezMes.value = 1 Then
        sUnaVezMes = "1"
    Else
        sUnaVezMes = "0"
    End If
    
    If chkUnaVezPeriodo.value = 1 Then
        sUnaVezPeriodo = "1"
    Else
        sUnaVezPeriodo = "0"
    End If
    
    
    If chkActivo.value = 1 Then
        sActivo = "1"
    Else
        sActivo = "0"
    End If
    
    If chkRRHH.value = 1 Then
        schkRRHH = "1"
    Else
        schkRRHH = "0"
    End If
    

    If chkMaximoUno.value = 1 Then
        sMaximoUno = "1"
    Else
        sMaximoUno = "0"
    End If
    
    If chkCalcVacacional.value = 1 And (chkSaldoInicialPeriodo.value = 1 Or chkSaldoSalidaEmpleado.value = 1) Then
        lbok = Mensaje("Se ha indicado cálculo vacacional pero también como Saldo de Período y/o Saldo x Salida Empleado. Son Excluyentes !!!", ICO_ERROR, False)
        Exit Sub
        
    End If
 
If bAdd Then
    If Not (rst.EOF And rst.BOF) Then
        sFiltro = "Tipo_Accion = '" & txtTipoAccion.Text + "'"
        If ExiteRstKey(rst, sFiltro) Then
           lbok = Mensaje("Ya existe esa Acción ", ICO_ERROR, False)
            txtTipoAccion.SetFocus
        Exit Sub
        End If
    End If

    If Not (rst.EOF And rst.BOF) Then
        sFiltro = "Prioridad = " & txtPrioridad.Text
        If ExiteRstKey(rst, sFiltro) Then
           lbok = Mensaje("Ya existe esa Prioridad", ICO_ERROR, False)
           txtPrioridad.SetFocus
        Exit Sub
        End If
    End If
    
    If Not (rst.EOF And rst.BOF) Then
        If schkRecalculoVac = "1" Then
            stmpValorSINO = "SI"
        Else
            stmpValorSINO = "NO"
        End If
        If stmpValorSINO = "SI" Then
            sFiltro = "CalculoVacacional = 'SI'"
            If ExiteRstKey(rst, sFiltro) Then
               lbok = Mensaje("Ya existe una Acción para Recálculo Vacacional", ICO_ERROR, False)
               ' txtTipoAccion.SetFocus
            Exit Sub
            End If
        End If
    End If


    If Not (rst.EOF And rst.BOF) Then
        If sEsSaldoInicialPeriodo = "1" Then
            stmpValorSINO = "SI"
        Else
            stmpValorSINO = "NO"
        End If
        If stmpValorSINO = "SI" Then
            sFiltro = "SaldoInicialPeriodo = 'SI'"
            If ExiteRstKey(rst, sFiltro) Then
               lbok = Mensaje("Ya existe una Acción para Saldo Inicial Período", ICO_ERROR, False)
               ' txtTipoAccion.SetFocus
            Exit Sub
            End If
        End If
    End If
    
    
    If Not (rst.EOF And rst.BOF) Then
        If schkAjuste = "1" Then
            stmpValorSINO = "SI"
        Else
            stmpValorSINO = "NO"
        End If
        If stmpValorSINO = "SI" Then
            If optResta.value = True Then
                sFiltro = "EsAjuste = 'SI' and Factor='RESTA'"
                If ExiteRstKey(rst, sFiltro) Then
                   lbok = Mensaje("Ya existe una Acción para Ajustes que RESTA dias", ICO_ERROR, False)
                   ' txtTipoAccion.SetFocus
               
                Exit Sub
                End If
            End If
            
            If optSuma.value = True Then
                sFiltro = "EsAjuste = 'SI' and Factor='SUMA'"
                If ExiteRstKey(rst, sFiltro) Then
                   lbok = Mensaje("Ya existe una Acción para Ajustes que SUMA dias", ICO_ERROR, False)
                   ' txtTipoAccion.SetFocus
               
                Exit Sub
                End If
            End If
            
        End If
              
    End If
    
    If Not (rst.EOF And rst.BOF) Then
        If sEsSaldoSalidaEmpleado = "1" Then
            stmpValorSINO = "SI"
        Else
            stmpValorSINO = "NO"
        End If
        If stmpValorSINO = "SI" Then
            sFiltro = "SaldoSalidaEmpleado = 'SI'"
            If ExiteRstKey(rst, sFiltro) Then
               lbok = Mensaje("Ya existe una Acción para Saldo por Salida de un Empleado", ICO_ERROR, False)
               ' txtTipoAccion.SetFocus
            Exit Sub
            End If
        End If
    End If
    
    
        lbok = sgvActualizaTipoAccion(txtTipoAccion.Text, txtDescr.Text, txtPrioridad.Text, sFactor, schkRRHH, _
        schkAfecta, schkRecalculoVac, schkAjuste, sEsSaldoInicialPeriodo, sSaldoSalidaEmpleado, sActivo, sSeRepiteMes, sUnaVezPeriodo, sUnaVezMes, sMaximoUno, "I")
        
        If lbok Then
            sMsg = "La Acción ha sido registrada exitosamente ... "
            lbok = Mensaje(sMsg, ICO_OK, False)
            ' actualiza datos
            cargaGridTiposAccion 2
            bEdit = False
            bAdd = False
            initControles
            IniciaIconos
        End If
bAdd = False
End If ' si estoy adicionando
If bEdit Then
    If Not (rst.EOF And rst.BOF) Then
        sFiltro = "Tipo_Accion<>'" & txtTipoAccion.Text & "' and Prioridad = " & txtPrioridad.Text
        If ExiteRstKey(rst, sFiltro) Then
           lbok = Mensaje("Ya existe esa Prioridad", ICO_ERROR, False)
           txtPrioridad.SetFocus
        Exit Sub
        End If
    End If
    
    
    If Not (rst.EOF And rst.BOF) Then
        If schkRecalculoVac = "1" Then
            stmpValorSINO = "SI"
        Else
            stmpValorSINO = "NO"
        End If
        If schkRecalculoVac = "1" Then
            sFiltro = "Tipo_Accion<>'" & txtTipoAccion.Text & "' and CalculoVacacional = 'SI'"
            If ExiteRstKey(rst, sFiltro) Then
               lbok = Mensaje("Ya existe esa Acción para Recálculo Vacacional", ICO_ERROR, False)
               ' txtTipoAccion.SetFocus
            Exit Sub
            End If
        End If
    End If

    If Not (rst.EOF And rst.BOF) Then

        If schkAjuste = "1" And Me.optSuma.value = True Then
            sFiltro = "Tipo_Accion<>'" & txtTipoAccion.Text & "' and EsAjuste = 'SI' and Factor = 'SUMA'"
            If ExiteRstKey(rst, sFiltro) Then
               lbok = Mensaje("Ya existe esa Acción para Ajustes que SUMAN dias", ICO_ERROR, False)
               ' txtTipoAccion.SetFocus
            Exit Sub
            End If
        End If
        
        If schkAjuste = "1" And Me.optResta.value = True Then
            sFiltro = "Tipo_Accion<>'" & txtTipoAccion.Text & "' and EsAjuste = 'SI' and Factor = 'RESTA'"
            If ExiteRstKey(rst, sFiltro) Then
               lbok = Mensaje("Ya existe esa Acción para Ajustes que RESTAN dias", ICO_ERROR, False)
               ' txtTipoAccion.SetFocus
            Exit Sub
            End If
        End If
        
    End If



    If Not (rst.EOF And rst.BOF) Then
        If sEsSaldoInicialPeriodo = "1" Then
            stmpValorSINO = "SI"
        Else
            stmpValorSINO = "NO"
        End If
        If sEsSaldoInicialPeriodo = "1" Then
            sFiltro = "Tipo_Accion<>'" & txtTipoAccion.Text & "' and SaldoInicialPeriodo = 'SI'"
            If ExiteRstKey(rst, sFiltro) Then
               lbok = Mensaje("Ya existe esa Acción para Saldo Inicial Periodo", ICO_ERROR, False)
               ' txtTipoAccion.SetFocus
            Exit Sub
            End If
        End If
    End If

    If Not (rst.EOF And rst.BOF) Then
        If sSaldoSalidaEmpleado = "1" Then
            stmpValorSINO = "SI"
        Else
            stmpValorSINO = "NO"
        End If
        If sSaldoSalidaEmpleado = "1" Then
            sFiltro = "Tipo_Accion<>'" & txtTipoAccion.Text & "' and SaldoSalidaEmpleado = 'SI'"
            If ExiteRstKey(rst, sFiltro) Then
               lbok = Mensaje("Ya existe esa Acción para Saldo Salida Empleado", ICO_ERROR, False)
               ' txtTipoAccion.SetFocus
            Exit Sub
            End If
        End If
    End If


        lbok = sgvActualizaTipoAccion(txtTipoAccion.Text, txtDescr.Text, txtPrioridad.Text, sFactor, schkRRHH, _
        schkAfecta, schkRecalculoVac, schkAjuste, sEsSaldoInicialPeriodo, sSaldoSalidaEmpleado, sActivo, sSeRepiteMes, sUnaVezPeriodo, sUnaVezMes, sMaximoUno, "E")
        
        If lbok Then
            sMsg = "Los datos fueron grabados Exitosamente ... "
            lbok = Mensaje(sMsg, ICO_OK, False)
            ' actualiza datos
            cargaGridTiposAccion 2
            bEdit = False
            bAdd = False
            initControles
            IniciaIconos
        End If
bEdit = False
End If ' si estoy adicionando
End Sub

Private Sub cmdVendB_Click()
If chkIndependiente.value = 0 Then
   Exit Sub
End If
Dim frm As frmBrowseCat
Set frm = New frmBrowseCat
frm.gsCaptionfrm = "Vendedores" '& lblund.Caption
frm.gsTablabrw = "vcomVendedores"
frm.gsCodigobrw = "CodVendedor"
frm.gsDescrbrw = "Nombre"
frm.gbFiltra = True
frm.gsFiltro = "Activo = 1 and CODSUCURSAL='" & sCodSucursal & "'"
frm.Show vbModal
If frm.gsCodigobrw <> "" Then
  txtVendedor.Text = frm.gsCodigobrw
End If

'If frm.gsDescrbrw <> "" Then
'  Me.lblNombre.Caption = frm.gsDescrbrw
'  ' fmtTextbox txtDescrSucursal, "R"
'  ' fmtTextbox txtCodSucursal, "O"
'End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdUndo_Click()
GetDataFromGridToControl
'sCodSucursal = txtCodSucursal.Text
IniciaIconos
End Sub

Private Sub Form_Load()
Set rst = New ADODB.Recordset
If rst.State = adStateOpen Then rst.Close
rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
rst.CursorLocation = adUseClient ' Cursor local al cliente
rst.LockType = adLockOptimistic
bEdit = False
bAdd = False
initControles
sSoloActivo = "1"
cargaGridTiposAccion sSoloActivo
End Sub

Private Sub IniciaIconos()
cmdSave.Enabled = False
cmdEditItem.Enabled = True
cmdEliminar.Enabled = True
cmdAdd.Enabled = True
bEdit = False
bAdd = False

End Sub
Private Sub cargaGridTiposAccion(sSoloActivos As String)
Dim sIndependiente As String
If rst.State = adStateOpen Then rst.Close
rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
rst.CursorLocation = adUseClient ' Cursor local al cliente
rst.LockType = adLockOptimistic
GSSQL = gsCompania & ".sgvGetTiposAccion" & " " & sSoloActivos
If rst.State = adStateOpen Then rst.Close
Set rst = GetRecordset(GSSQL)
If Not (rst.EOF And rst.BOF) Then
  Set TDBG.DataSource = rst
  'CargarDatos rst, TDBG, "Codigo", "Descr"
  TDBG.Refresh
  IniciaIconos
End If
End Sub

Private Sub optallSuc_Click()
cmdSucursal.Enabled = False
txtCodSucursal.Text = ""
txtCodSucursal.Enabled = False
txtDescrSucursal.Text = ""
txtDescrSucursal.Enabled = False
End Sub

Private Sub optunasuc_Click()
cmdSucursal.Enabled = True
txtCodSucursal.Text = ""
txtCodSucursal.Enabled = True
txtDescrSucursal.Text = ""
txtDescrSucursal.Enabled = True
End Sub

Private Sub TDBG_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
GetDataFromGridToControl
'sCodSucursal = txtCodSucursal.Text
IniciaIconos
End Sub


