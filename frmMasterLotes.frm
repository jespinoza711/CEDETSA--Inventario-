VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMasterLotes 
   BackColor       =   &H00FEE3DA&
   Caption         =   "Form1"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10275
   FillColor       =   &H00F4D5BB&
   Icon            =   "frmMasterLotes.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6690
   ScaleWidth      =   10275
   Begin VB.Frame Frame 
      BackColor       =   &H00FEE3DA&
      Height          =   795
      Left            =   120
      TabIndex        =   13
      Top             =   450
      Width           =   9915
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
         Left            =   2460
         Picture         =   "frmMasterLotes.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Deshacer / Cancelar"
         Top             =   150
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
         Left            =   1260
         Picture         =   "frmMasterLotes.frx":1594
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Elimina el item actualmente seleccionado en el grid de datos ..."
         Top             =   150
         Width           =   555
      End
      Begin VB.CommandButton cmdSave 
         Enabled         =   0   'False
         Height          =   555
         Left            =   1860
         Picture         =   "frmMasterLotes.frx":225E
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Aplica y Guarda los datos de la transacción en Firme ..."
         Top             =   150
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
         Left            =   660
         Picture         =   "frmMasterLotes.frx":3F28
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Modifica los datos mostrados en el Grid con los datos digitados ..."
         Top             =   150
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
         Left            =   60
         Picture         =   "frmMasterLotes.frx":4BF2
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Agrega el item con los datos digitados..."
         Top             =   150
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FEE3DA&
      Height          =   1905
      Left            =   120
      TabIndex        =   0
      Top             =   1230
      Width           =   9930
      Begin MSComCtl2.DTPicker dtpFechaVencimiento 
         Height          =   375
         Left            =   2220
         TabIndex        =   11
         Top             =   1140
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61472769
         CurrentDate     =   41772
      End
      Begin VB.TextBox txtLoteProveedor 
         Appearance      =   0  'Flat
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
         Left            =   7050
         TabIndex        =   7
         Top             =   690
         Width           =   2655
      End
      Begin VB.TextBox txtLoteInterno 
         Appearance      =   0  'Flat
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
         TabIndex        =   2
         Top             =   690
         Width           =   2595
      End
      Begin VB.TextBox txtIDLote 
         Appearance      =   0  'Flat
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
         TabIndex        =   1
         Top             =   300
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtpFechaProduccion 
         Height          =   375
         Left            =   8310
         TabIndex        =   12
         Top             =   1170
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   61472769
         CurrentDate     =   41772
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Producción:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6120
         TabIndex        =   10
         Top             =   1230
         Width           =   1965
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Vencimiento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   1230
         Width           =   1965
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote Proveedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5460
         TabIndex        =   8
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote Interno:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ID Lote:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   795
      End
   End
   Begin TrueOleDBGrid60.TDBGrid TDBG 
      Height          =   3315
      Left            =   120
      OleObjectBlob   =   "frmMasterLotes.frx":58BC
      TabIndex        =   5
      Top             =   3270
      Width           =   9945
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
      Left            =   -420
      TabIndex        =   6
      Top             =   -30
      Width           =   11070
   End
   Begin VB.Image Image1 
      Height          =   885
      Left            =   -330
      Picture         =   "frmMasterLotes.frx":B34D
      Stretch         =   -1  'True
      Top             =   -450
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
Dim lbOk As Boolean
lbOk = False
On Error GoTo error

    If ExisteDependencia("invEXISTENCIALOTE", sFldname, sFldVal, "N") Then
        lbOk = True
        GoTo Salir
    Else
        If ExisteDependencia("invMovimientos", sFldname, sFldVal, "N") Then
            lbOk = True
            GoTo Salir
        End If
    End If

Salir:
DependenciaLote = lbOk
Exit Function

error:
    lbOk = False
    GoTo Salir
End Function


Private Sub cmdEliminar_Click()
    Dim lbOk As Boolean
    Dim sMsg As String
    Dim sTipo As String
    Dim sFiltro As String
    Dim sActivo As String
    Dim sFactura As String

    If txtIDLote.Text = "" Then
        lbOk = Mensaje("El IDLote no puede estar en Blanco", ICO_ERROR, False)
        Exit Sub
    End If
    
    ' hay que validar la integridad referencial
     'Validar la dependecia de la bodega
    If DependenciaLote("IDLote", rst!IdLote) Then
        lbOk = Mensaje("No se puede eliminar, el Lote tiene Asociada transacciones", ICO_ERROR, False)
        Exit Sub
    End If
    
    lbOk = Mensaje("Está seguro de eliminar el Lote " & rst("IDLote").value, ICO_PREGUNTA, True)
    If lbOk Then
                lbOk = invUpdateLote("D", txtIDLote.Text, txtLoteInterno.Text, txtLoteProveedor.Text, dtpFechaVencimiento.value, dtpFechaProduccion.value)
        
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
    Dim sFactura As String
    Dim sFiltro As String
    If txtIDLote.Text = "" Then
        lbOk = Mensaje("IDLote no puede estar en Blanco", ICO_ERROR, False)
        Exit Sub
    End If
    
    If txtLoteInterno.Text = "" Then
        lbOk = Mensaje("El Lote interno no puede estar en blanco", ICO_ERROR, False)
        Exit Sub
    End If
    
    If txtLoteProveedor.Text = "" Then
        lbOk = Mensaje("El Lote proveedor no puede estar en blanco", ICO_ERROR, False)
        Exit Sub
    End If
    
    
    If (dtpFechaVencimiento.value = dtpFechaProduccion.value) Then
        lbOk = Mensaje("La fecha de vencimiento y fecha de producción no pueden ser iguales", ICO_ERROR, False)
        Exit Sub
    End If
        
    If Accion = Add Then
    
        If Not (rst.EOF And rst.BOF) Then
            sFiltro = "IDLote = '" & txtIDLote.Text & "'"
            If ExiteRstKey(rst, sFiltro) Then
               lbOk = Mensaje("Ya exista el Lote ", ICO_ERROR, False)
                txtIDLote.SetFocus
            Exit Sub
            End If
        End If
    
            lbOk = invUpdateLote("I", txtIDLote.Text, txtLoteInterno.Text, txtLoteProveedor.Text, dtpFechaVencimiento.value, dtpFechaProduccion.value)
            
            If lbOk Then
                sMsg = "El Lote ha sido registrada exitosamente ... "
                lbOk = Mensaje(sMsg, ICO_OK, False)
                ' actualiza datos
                cargaGrid
                Accion = View
                HabilitarControles
                HabilitarBotones
            Else
                 sMsg = "Ha ocurrido un error tratando de Agregar el Lote... "
                lbOk = Mensaje(sMsg, ICO_ERROR, False)
            End If
    End If ' si estoy adicionando
        If Accion = Edit Then
            If Not (rst.EOF And rst.BOF) Then
                lbOk = invUpdateLote("U", txtIDLote.Text, txtLoteInterno.Text, txtLoteProveedor.Text, dtpFechaVencimiento.value, dtpFechaProduccion.value)
                If lbOk Then
                    sMsg = "El Lote ha sido registrada exitosamente ..."
                    lbOk = Mensaje(sMsg, ICO_OK, False)
                    ' actualiza datos
                    cargaGrid
                    Accion = View
                    HabilitarControles
                    HabilitarBotones
                Else
                    sMsg = "Ha ocurrido un error tratando de Agregar el Lote... "
                    lbOk = Mensaje(sMsg, ICO_ERROR, False)
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

Private Sub Form_Activate()
    HighlightInWin Me.Name
    SetupFormToolbar (Me.Name)
End Sub

Private Sub Form_Load()
    MDIMain.AddForm Me.Name
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
    SetupFormToolbar ("no name")
    'Main.SubtractForm Me.Name
    Set frmBodega = Nothing
End Sub


