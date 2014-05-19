VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmTablas 
   Caption         =   "Tablas Globales del Sistema"
   ClientHeight    =   6720
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   9465
   StartUpPosition =   1  'CenterOwner
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
      Height          =   555
      Left            =   8565
      Picture         =   "frmTablas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Deshacer / Cancelar"
      Top             =   4680
      Width           =   555
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
      Height          =   555
      Left            =   8565
      Picture         =   "frmTablas.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Elimina el item actualmente seleccionado en el grid de datos ..."
      Top             =   3480
      Width           =   555
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
      Height          =   555
      Left            =   8565
      Picture         =   "frmTablas.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Agrega el item con los datos digitados..."
      Top             =   2280
      Width           =   555
   End
   Begin VB.CommandButton cmdSave 
      Enabled         =   0   'False
      Height          =   555
      Left            =   8565
      Picture         =   "frmTablas.frx":265E
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Aplica y Guarda los datos de la transacción en Firme ..."
      Top             =   4080
      Width           =   555
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
      Height          =   555
      Left            =   8565
      Picture         =   "frmTablas.frx":4328
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Modifica los datos mostrados en el Grid con los datos digitados ..."
      Top             =   2880
      Width           =   555
   End
   Begin VB.Frame Frame2 
      Height          =   1395
      Left            =   165
      TabIndex        =   2
      Top             =   720
      Width           =   9135
      Begin VB.CommandButton cmdModulo 
         Height          =   320
         Left            =   960
         Picture         =   "frmTablas.frx":4FF2
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   300
      End
      Begin VB.TextBox txtDecrModulo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002F2F2F&
         Height          =   285
         Left            =   3720
         TabIndex        =   12
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtIDTabla 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002F2F2F&
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelModulo 
         Height          =   320
         Left            =   2520
         Picture         =   "frmTablas.frx":5334
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   300
      End
      Begin VB.TextBox txtAbrev 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002F2F2F&
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002F2F2F&
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   600
         Width           =   1095
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
         ForeColor       =   &H002F2F2F&
         Height          =   285
         Left            =   3720
         TabIndex        =   1
         Top             =   600
         Width           =   5055
      End
      Begin VB.CheckBox chkActivo 
         Caption         =   "Activo ?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002F2F2F&
         Height          =   255
         Left            =   7800
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Descr :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002F2F2F&
         Height          =   255
         Left            =   2880
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Módulo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002F2F2F&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Abreviatura:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002F2F2F&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Código :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002F2F2F&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H002F2F2F&
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
   End
   Begin TrueOleDBGrid60.TDBGrid TDBG 
      Height          =   4365
      Left            =   150
      OleObjectBlob   =   "frmTablas.frx":5776
      TabIndex        =   6
      Top             =   2280
      Width           =   8205
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tablas Globales del Sistema"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002F2F2F&
      Height          =   375
      Left            =   -180
      TabIndex        =   15
      Top             =   -30
      Width           =   10140
   End
   Begin VB.Image Image1 
      Height          =   885
      Left            =   -435
      Picture         =   "frmTablas.frx":A86F
      Stretch         =   -1  'True
      Top             =   -345
      Width           =   11490
   End
End
Attribute VB_Name = "frmTablas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As ADODB.Recordset
Dim bOrdenCodigo As Boolean
Dim bOrdenDescr As Boolean
Dim sCodSucursal As String
Dim Accion As TypAccion
Dim sSoloActivo As String

Private Sub HabilitarBotones()
    Select Case Accion
        Case TypAccion.Add, TypAccion.Edit
            cmdSave.Enabled = True
            cmdUndo.Enabled = True
            cmdEliminar.Enabled = False
            cmdAdd.Enabled = False
            cmdEditItem.Enabled = False
        Case TypAccion.View
            cmdSave.Enabled = False
            cmdUndo.Enabled = False
            cmdEliminar.Enabled = True
            cmdAdd.Enabled = True
            cmdEditItem.Enabled = True
    End Select
End Sub

Public Sub HabilitarControles()
    Select Case Accion
        Case TypAccion.Add
            txtCodigo.Enabled = True
            txtDescr.Enabled = True
            chkActivo.Enabled = True
            chkActivo.value = 1
            txtCodigo.Text = ""
            txtDescr.Text = ""
            fmtTextbox txtCodigo, "O"
            fmtTextbox txtDescr, "O"
        Case TypAccion.Edit
            txtCodigo.Enabled = True
            txtDescr.Enabled = True
            fmtTextbox txtCodigo, "R"
            fmtTextbox txtDescr, "O"
            chkActivo.Enabled = True
        Case TypAccion.View
            'txtVendedor.Enabled = False
            'txtNombre.Enabled = False
            chkActivo.Enabled = False
            'txtTipo.Enabled = False
            cmdTipo.Enabled = False
            fmtTextbox Me.txtCodigo, "R"
            fmtTextbox txtDescr, "R"
            Me.chkActivo.Enabled = False
    End Select
End Sub

Private Sub cmdAdd_Click()
    Accion = Add
    HabilitarBotones
    HabilitarControles
    txtCodigo.SetFocus
End Sub

Private Sub cmdEditItem_Click()
    Accion = Edit
    GetDataFromGridToControl
    HabilitarBotones
    HabilitarControles
End Sub
Private Sub GetDataFromGridToControl()
If Not (rst.EOF And rst.BOF) Then
    txtCodigo.Text = rst("Departamento").value
    txtDescr.Text = rst("Descr").value
    If rst("Activo").value = True Then
        chkActivo.value = 1
    Else
        chkActivo.value = 0
    End If
Else
    txtCodigo.Text = ""
    txtDescr.Text = ""
End If

End Sub

Private Sub cmdEliminar_Click()
Dim lbOk As Boolean
Dim sMsg As String
Dim sTipo As String
Dim sFiltro As String
Dim sActivo As String

    If txtCodigo.Text = "" Then
        lbOk = Mensaje("El Departamento no puede estar en Blanco", ICO_ERROR, False)
        Exit Sub
    End If
    If chkActivo.value = 1 Then
        sActivo = "1"
    Else
        sActivo = "0"
    End If
    ' hay que validar la integridad referencial
    lbOk = Mensaje("Está seguro de eliminar el Centro de Costo " & rst("Departamento").value, ICO_PREGUNTA, True)
    If lbOk Then
               ' lbOk = sgvActualizaDepartamento(txtCodigo.Text, txtDescr.Text, sActivo, "D")
        
        If lbOk Then
            sMsg = "Borrado Exitosamente ... "
            lbOk = Mensaje(sMsg, ICO_OK, False)
            ' actualiza datos
            cargaGrid
        End If
    End If
End Sub

Private Sub cmdSave_Click()
Dim lbOk As Boolean
Dim sMsg As String
Dim sActivo As String
Dim sFiltro As String
    If txtCodigo.Text = "" Then
        lbOk = Mensaje("El Departamento no pueden estar en Blanco", ICO_ERROR, False)
        Exit Sub
    End If
    If Me.chkActivo.value = 1 Then
        sActivo = "1"
    Else
        sActivo = "0"
    End If
    
    If txtDescr.Text = "" Then
        lbOk = Mensaje("La Descripción del Departamento no puede estar en blanco", ICO_ERROR, False)
        Exit Sub
    End If
    

        
If bAdd Then

    If Not (rst.EOF And rst.BOF) Then
        sFiltro = "Departamento = '" & txtCodigo.Text & "'"
        If ExiteRstKey(rst, sFiltro) Then
           lbOk = Mensaje("Ya existe ese Departamento ", ICO_ERROR, False)
            txtCodigo.SetFocus
        Exit Sub
        End If
    End If

        'lbOk = sgvActualizaDepartamento(txtCodigo.Text, txtDescr.Text, sActivo, "I")
        
        If lbOk Then
            sMsg = "El Departamento ha sido registrado exitosamente ... "
            lbOk = Mensaje(sMsg, ICO_OK, False)
            ' actualiza datos
            cargaGrid
            bEdit = False
            bAdd = False
            'initControles
          '  IniciaIconos
        End If
bAdd = False
End If ' si estoy adicionando
If bEdit Then
    If Not (rst.EOF And rst.BOF) Then
      ' lbOk = sgvActualizaDepartamento(txtCodigo.Text, txtDescr.Text, sActivo, "E")
        If lbOk Then
            sMsg = "Los datos fueron grabados Exitosamente ... "
            lbOk = Mensaje(sMsg, ICO_OK, False)
            ' actualiza datos
            cargaGrid
            bEdit = False
            bAdd = False
'            initControles
'            IniciaIconos
        End If
    End If
bEdit = False
End If ' si estoy adicionando

End Sub

Private Sub cmdUndo_Click()
    Accion = View
    GetDataFromGridToControl
    HabilitarControles
    HabilitarBotones
End Sub

Private Sub Form_Load()
    Set rst = New ADODB.Recordset
    If rst.State = adStateOpen Then rst.Close
    rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
    rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
    rst.CursorLocation = adUseClient ' Cursor local al cliente
    rst.LockType = adLockOptimistic
    Accion = View
    HabilitarControles
    HabilitarBotones
    cargaGrid
End Sub



Private Sub cargaGrid()
    Dim sIndependiente As String
    If rst.State = adStateOpen Then rst.Close
    rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
    rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
    rst.CursorLocation = adUseClient ' Cursor local al cliente
    rst.LockType = adLockOptimistic
    GSSQL = gsCompania & ".globalGetTablas -1"
    If rst.State = adStateOpen Then rst.Close
    Set rst = GetRecordset(GSSQL)
    If Not (rst.EOF And rst.BOF) Then
      Set TDBG.DataSource = rst
      'CargarDatos rst, TDBG, "Codigo", "Descr"
      TDBG.Refresh
      'IniciaIconos
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not (rst Is Nothing) Then Set rst = Nothing
End Sub

Private Sub TDBG_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    GetDataFromGridToControl
'sCodSucursal = txtCodSucursal.Text
    HabilitarControles
    HabilitarBotones
End Sub


