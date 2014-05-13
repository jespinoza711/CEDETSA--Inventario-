VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMasterLotes 
   Caption         =   "Form1"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1905
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   8940
      Begin MSComCtl2.DTPicker dtpFechaVencimiento 
         Height          =   375
         Left            =   2220
         TabIndex        =   16
         Top             =   1140
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20971521
         CurrentDate     =   41772
      End
      Begin VB.TextBox txtLoteProveedor 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6030
         TabIndex        =   12
         Top             =   690
         Width           =   2655
      End
      Begin VB.TextBox txtLoteInterno 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1380
         TabIndex        =   7
         Top             =   690
         Width           =   2595
      End
      Begin VB.TextBox txtIDLote 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1380
         TabIndex        =   6
         Top             =   300
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtpFechaProduccion 
         Height          =   375
         Left            =   7260
         TabIndex        =   17
         Top             =   1170
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20971521
         CurrentDate     =   41772
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha de Producción:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5100
         TabIndex        =   15
         Top             =   1230
         Width           =   1965
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha de Vencimiento:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   1230
         Width           =   1965
      End
      Begin VB.Label Label2 
         Caption         =   "Lote Proveedor:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4500
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Lote Interno:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "ID Lote:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   795
      End
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
      Left            =   8355
      Picture         =   "frmMasterLotes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Modifica los datos mostrados en el Grid con los datos digitados ..."
      Top             =   3510
      Width           =   555
   End
   Begin VB.CommandButton cmdSave 
      Enabled         =   0   'False
      Height          =   555
      Left            =   8355
      Picture         =   "frmMasterLotes.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Aplica y Guarda los datos de la transacción en Firme ..."
      Top             =   4710
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
      Left            =   8355
      Picture         =   "frmMasterLotes.frx":2994
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Agrega el item con los datos digitados..."
      Top             =   2910
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
      Left            =   8355
      Picture         =   "frmMasterLotes.frx":365E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Elimina el item actualmente seleccionado en el grid de datos ..."
      Top             =   4110
      Width           =   555
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
      Height          =   555
      Left            =   8355
      Picture         =   "frmMasterLotes.frx":4328
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Deshacer / Cancelar"
      Top             =   5310
      Width           =   555
   End
   Begin TrueOleDBGrid60.TDBGrid TDBG 
      Height          =   3135
      Left            =   210
      OleObjectBlob   =   "frmMasterLotes.frx":4FF2
      TabIndex        =   10
      Top             =   2850
      Width           =   7965
   End
   Begin VB.Label lbFormCaption 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Titulo Catalogo"
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
      Left            =   -390
      TabIndex        =   11
      Top             =   -15
      Width           =   10140
   End
   Begin VB.Image Image1 
      Height          =   885
      Left            =   -330
      Picture         =   "frmMasterLotes.frx":AA73
      Stretch         =   -1  'True
      Top             =   -330
      Width           =   11490
   End
End
Attribute VB_Name = "frmMasterLotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rst As ADODB.Recordset
Dim bOrdenCodigo As Boolean
Dim bOrdenDescr As Boolean
Dim sCodSucursal As String
Dim sSoloActivo As String
Dim Accion As TypAccion
Public gsFormCaption As String
Public gsTitle As String

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
            txtIDLote.Enabled = True
            txtLoteInterno.Enabled = True
            txtLoteProveedor.Enabled = True
            txtIDLote.Text = "100"
            txtLoteInterno.Text = ""
            txtLoteProveedor.Text = ""
            fmtTextbox txtIDLote, "R"
            fmtTextbox txtLoteInterno, "O"
            fmtTextbox txtLoteProveedor, "O"
            dtpFechaVencimiento.Enabled = True
            dtpFechaProduccion.Enabled = True
            Me.TDBG.Enabled = False
        Case TypAccion.Edit
            txtLoteInterno.Enabled = True
            fmtTextbox txtIDLote, "R"
            fmtTextbox txtLoteInterno, "O"
            fmtTextbox txtLoteProveedor, "O"
            dtpFechaVencimiento.Enabled = True
            dtpFechaProduccion.Enabled = True
            Me.TDBG.Enabled = False
        Case TypAccion.View
            fmtTextbox txtLoteInterno, "R"
            fmtTextbox txtIDLote, "R"
            fmtTextbox txtLoteProveedor, "R"
            dtpFechaVencimiento.Enabled = False
            dtpFechaProduccion.Enabled = False
            Me.TDBG.Enabled = True
    End Select
End Sub

Private Sub cmdAdd_Click()
    Accion = Add
    HabilitarBotones
    HabilitarControles
    txtLoteInterno.SetFocus
End Sub

Private Sub cmdEditItem_Click()
    Accion = Edit
    GetDataFromGridToControl
    HabilitarBotones
    HabilitarControles
End Sub
Private Sub GetDataFromGridToControl()
If Not (rst.EOF And rst.BOF) Then
    txtIDLote.Text = rst("IDLote").value
    txtLoteInterno.Text = rst("LoteInterno").value
    txtLoteProveedor.Text = rst("LoteProveedor").value
    dtpFechaVencimiento.value = rst("FechaVencimiento").value
    dtpFechaProduccion.value = rst("FechaFabricacion").value
Else
    txtIDLote.Text = ""
    txtLoteInterno.Text = ""
    txtLoteProveedor.Text = ""
    dtpFechaVencimiento.value = DateTime.Now
    dtpFechaProduccion.value = DateTime.Now
End If

End Sub


Public Function DependenciaLote(sFldname As String, sFldVal As String) As Boolean
Dim lbok As Boolean
lbok = False
On Error GoTo error

    If ExisteDependencia("invEXISTENCIALOTE", sFldname, sFldVal, "N") Then
        lbok = True
        GoTo Salir
    Else
        If ExisteDependencia("invMovimientos", sFldname, sFldVal, "N") Then
            lbok = True
            GoTo Salir
        End If
    End If

Salir:
DependenciaLote = lbok
Exit Function

error:
    lbok = False
    GoTo Salir
End Function


Private Sub cmdEliminar_Click()
    Dim lbok As Boolean
    Dim sMsg As String
    Dim sTipo As String
    Dim sFiltro As String
    Dim sActivo As String
    Dim sFactura As String

    If txtIDLote.Text = "" Then
        lbok = Mensaje("El IDLote no puede estar en Blanco", ICO_ERROR, False)
        Exit Sub
    End If
    
    ' hay que validar la integridad referencial
     'Validar la dependecia de la bodega
    If DependenciaLote("IDLote", rst!IDLote) Then
        lbok = Mensaje("No se puede eliminar, el Lote tiene Asociada transacciones", ICO_ERROR, False)
        Exit Sub
    End If
    
    lbok = Mensaje("Está seguro de eliminar el Lote " & rst("IDLote").value, ICO_PREGUNTA, True)
    If lbok Then
                lbok = invUpdateLote("D", txtIDLote.Text, txtLoteInterno.Text, txtLoteProveedor.Text, dtpFechaVencimiento.value, dtpFechaProduccion.value)
        
        If lbok Then
            sMsg = "Borrado Exitosamente ... "
            lbok = Mensaje(sMsg, ICO_OK, False)
            ' actualiza datos
            cargaGrid
        End If
    End If
End Sub

Private Sub cmdSave_Click()
    Dim lbok As Boolean
    Dim sMsg As String
    Dim sActivo As String
    Dim sFactura As String
    Dim sFiltro As String
    If txtIDLote.Text = "" Then
        lbok = Mensaje("IDLote no puede estar en Blanco", ICO_ERROR, False)
        Exit Sub
    End If
    
    If txtLoteInterno.Text = "" Then
        lbok = Mensaje("El Lote interno no puede estar en blanco", ICO_ERROR, False)
        Exit Sub
    End If
    
    If txtLoteProveedor.Text = "" Then
        lbok = Mensaje("El Lote proveedor no puede estar en blanco", ICO_ERROR, False)
        Exit Sub
    End If
    
    
    If (dtpFechaVencimiento.value = dtpFechaProduccion.value) Then
        lbok = Mensaje("La fecha de vencimiento y fecha de producción no pueden ser iguales", ICO_ERROR, False)
        Exit Sub
    End If
        
    If Accion = Add Then
    
        If Not (rst.EOF And rst.BOF) Then
            sFiltro = "IDLote = '" & txtIDLote.Text & "'"
            If ExiteRstKey(rst, sFiltro) Then
               lbok = Mensaje("Ya exista el Lote ", ICO_ERROR, False)
                txtIDLote.SetFocus
            Exit Sub
            End If
        End If
    
            lbok = invUpdateLote("I", txtIDLote.Text, txtLoteInterno.Text, txtLoteProveedor.Text, dtpFechaVencimiento.value, dtpFechaProduccion.value)
            
            If lbok Then
                sMsg = "El Lote ha sido registrada exitosamente ... "
                lbok = Mensaje(sMsg, ICO_OK, False)
                ' actualiza datos
                cargaGrid
                Accion = View
                HabilitarControles
                HabilitarBotones
            Else
                 sMsg = "Ha ocurrido un error tratando de Agregar el Lote... "
                lbok = Mensaje(sMsg, ICO_ERROR, False)
            End If
    End If ' si estoy adicionando
        If Accion = Edit Then
            If Not (rst.EOF And rst.BOF) Then
                lbok = invUpdateLote("U", txtIDLote.Text, txtLoteInterno.Text, txtLoteProveedor.Text, dtpFechaVencimiento.value, dtpFechaProduccion.value)
                If lbok Then
                    sMsg = "El Lote ha sido registrada exitosamente ..."
                    lbok = Mensaje(sMsg, ICO_OK, False)
                    ' actualiza datos
                    cargaGrid
                    Accion = View
                    HabilitarControles
                    HabilitarBotones
                Else
                    sMsg = "Ha ocurrido un error tratando de Agregar el Lote... "
                    lbok = Mensaje(sMsg, ICO_ERROR, False)
                End If
            End If
        
    End If ' si estoy adicionando

End Sub

Private Sub cmdUndo_Click()
    GetDataFromGridToControl
    Accion = View
    HabilitarBotones
    HabilitarControles
End Sub

Private Sub Form_Load()
    Set rst = New ADODB.Recordset
    If rst.State = adStateOpen Then rst.Close
    rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
    rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
    rst.CursorLocation = adUseClient ' Cursor local al cliente
    rst.LockType = adLockOptimistic
    Me.Caption = gsFormCaption
    Me.lbFormCaption = gsTitle
    Accion = View
    HabilitarBotones
    HabilitarControles
    cargaGrid
End Sub



Private Sub cargaGrid()
    Dim sIndependiente As String
    If rst.State = adStateOpen Then rst.Close
    rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
    rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
    rst.CursorLocation = adUseClient ' Cursor local al cliente
    rst.LockType = adLockOptimistic
    GSSQL = gsCompania & ".invGetLotes -1"
    If rst.State = adStateOpen Then rst.Close
    Set rst = GetRecordset(GSSQL)
    If Not (rst.EOF And rst.BOF) Then
      Set TDBG.DataSource = rst
      'CargarDatos rst, TDBG, "Codigo", "Descr"
      TDBG.Refresh
      'IniciaIconos
    End If
End Sub


Private Sub TDBG_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    GetDataFromGridToControl
    HabilitarControles
    HabilitarBotones
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (rst Is Nothing) Then Set rst = Nothing
End Sub


