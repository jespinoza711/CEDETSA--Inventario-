VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMasterLotes 
   Caption         =   "Form1"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10425
   FillColor       =   &H00F4D5BB&
   Icon            =   "frmMasterLotes.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8130
   ScaleWidth      =   10425
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   10425
      TabIndex        =   18
      Top             =   0
      Width           =   10425
      Begin VB.Image Image 
         Height          =   480
         Index           =   2
         Left            =   150
         Picture         =   "frmMasterLotes.frx":08CA
         Top             =   45
         Width           =   480
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Actualización del Maestro de Lotes de Productos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   165
         Index           =   1
         Left            =   930
         TabIndex        =   20
         Top             =   420
         Width           =   2955
      End
      Begin VB.Label lbFormCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caption"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   930
         TabIndex        =   19
         Top             =   90
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1905
      Left            =   120
      TabIndex        =   0
      Top             =   1050
      Width           =   9990
      Begin MSComCtl2.DTPicker dtpFechaVencimiento 
         Height          =   315
         Left            =   2220
         TabIndex        =   10
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   20971521
         CurrentDate     =   41772
      End
      Begin VB.TextBox txtLoteProveedor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7050
         TabIndex        =   6
         Top             =   750
         Width           =   2655
      End
      Begin VB.TextBox txtLoteInterno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1380
         TabIndex        =   2
         Top             =   750
         Width           =   2595
      End
      Begin VB.TextBox txtIDLote 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         Height          =   315
         Left            =   8310
         TabIndex        =   11
         Top             =   1230
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
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
         Format          =   20971521
         CurrentDate     =   41772
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Producción:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   6270
         TabIndex        =   9
         Top             =   1290
         Width           =   1965
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Vencimiento:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   1290
         Width           =   1965
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote Proveedor:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   5460
         TabIndex        =   7
         Top             =   780
         Width           =   1395
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote Interno:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   780
         Width           =   1125
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ID Lote:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   795
      End
   End
   Begin TrueOleDBGrid60.TDBGrid TDBG 
      Height          =   4935
      Left            =   120
      OleObjectBlob   =   "frmMasterLotes.frx":1194
      TabIndex        =   5
      Top             =   3090
      Width           =   9945
   End
   Begin VB.Frame Frame 
      Height          =   795
      Left            =   120
      TabIndex        =   12
      Top             =   930
      Visible         =   0   'False
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
         Picture         =   "frmMasterLotes.frx":6C29
         Style           =   1  'Graphical
         TabIndex        =   17
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
         Picture         =   "frmMasterLotes.frx":78F3
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Elimina el item actualmente seleccionado en el grid de datos ..."
         Top             =   150
         Width           =   555
      End
      Begin VB.CommandButton cmdSave 
         Enabled         =   0   'False
         Height          =   555
         Left            =   1860
         Picture         =   "frmMasterLotes.frx":85BD
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Picture         =   "frmMasterLotes.frx":A287
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Picture         =   "frmMasterLotes.frx":AF51
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Agrega el item con los datos digitados..."
         Top             =   150
         Width           =   555
      End
   End
   Begin Inventario.CtlLiner CtlLiner 
      Height          =   30
      Left            =   0
      TabIndex        =   21
      Top             =   750
      Width           =   17925
      _ExtentX        =   31618
      _ExtentY        =   53
   End
End
Attribute VB_Name = "frmMasterLotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rst As ADODB.Recordset
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
    HabilitarBotonesMain
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
On Error GoTo ERROR

    If ExisteDependencia("invEXISTENCIALOTE", sFldname, sFldVal, "N") Then
        lbok = True
        GoTo salir
    Else
        If ExisteDependencia("invMovimientos", sFldname, sFldVal, "N") Then
            lbok = True
            GoTo salir
        End If
    End If

salir:
DependenciaLote = lbok
Exit Function

ERROR:
    lbok = False
    GoTo salir
End Function


Private Sub cmdEliminar_Click()
    Dim lbok As Boolean
    Dim sMsg As String
    
    If txtIDLote.Text = "" Then
        lbok = Mensaje("El IDLote no puede estar en Blanco", ICO_ERROR, False)
        Exit Sub
    End If
    
    ' hay que validar la integridad referencial
     'Validar la dependecia de la bodega
    If DependenciaLote("IDLote", rst!IdLote) Then
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
    SetupFormToolbar ("no form")
    MDIMain.SubtractForm Me.Name
    Set frmBodega = Nothing
End Sub

Public Sub CommandPass(ByVal srcPerformWhat As String)
    On Error GoTo err
    Select Case srcPerformWhat
        Case "Nuevo"
            cmdAdd_Click
        Case "Editar"
            cmdEditItem_Click
        Case "Eliminar"
            cmdEliminar_Click
        Case "Cancelar"
            cmdUndo_Click
        Case "Imprimir"
            MsgBox "Imprimir"
        Case "Exportar"
            ExportaGridToExcel Me.TDBG, "Listado de Lotes"
        Case "Cerrar"
            Unload Me
        Case "Guardar"
            cmdSave_Click
    End Select
    Exit Sub
    'Trap the error
err:
    If err.Number = -2147467259 Then
        MsgBox "You cannot delete this record because it was used by other records! If you want to delete this record" & vbCrLf & _
               "you will first have to delete or change the records that currenly used this record as shown bellow." & vbCrLf & vbCrLf & _
               err.Description, , "Delete Operation Failed!"
        Me.MousePointer = vbDefault
    End If
End Sub


Private Sub HabilitarBotonesMain()
    Select Case Accion
        Case TypAccion.Add, TypAccion.Edit
            MDIMain.tbMenu.Buttons(12).Enabled = True 'Guardar
            MDIMain.tbMenu.Buttons(13).Enabled = True 'Undo
            MDIMain.tbMenu.Buttons(11).Enabled = False 'Eliminar
            MDIMain.tbMenu.Buttons(8).Enabled = False 'Nuevo
            MDIMain.tbMenu.Buttons(10).Enabled = False 'Editar
        Case TypAccion.View
            MDIMain.tbMenu.Buttons(12).Enabled = False 'Guardar
            MDIMain.tbMenu.Buttons(13).Enabled = False 'Undo
            MDIMain.tbMenu.Buttons(11).Enabled = True 'Eliminar
            MDIMain.tbMenu.Buttons(8).Enabled = True 'Nuevo
            MDIMain.tbMenu.Buttons(10).Enabled = True 'Editar
    End Select
End Sub


Private Sub Form_Resize()
 On Error Resume Next
    If WindowState <> vbMinimized Then
        If Me.Width < 9195 Then Me.Width = 9195
        If Me.Height < 4500 Then Me.Height = 4500
        
        'center_obj_horizontal Me, Frame2
        'Frame2.Width = ScaleWidth - CONTROL_MARGIN
        
        TDBG.Width = Me.ScaleWidth - CONTROL_MARGIN
        TDBG.Height = (Me.ScaleHeight - Me.picHeader.Height) - TDBG.top
        
    End If
    TrueDBGridResize 1
End Sub

Public Sub TrueDBGridResize(iIndex As Integer)
    'If WindowState <> vbMaximized Then Exit Sub
    Dim i As Integer
    Dim dAnchoTotal As Double
    Dim dAnchocol As Double
    dAnchoTotal = 0
    dAnchocol = 0
    For i = 0 To Me.TDBG.Columns.Count - 1
        If (i = iIndex) Then
            dAnchocol = TDBG.Columns(i).Width
        Else
            dAnchoTotal = dAnchoTotal + TDBG.Columns(i).Width
        End If
    Next i

    Me.TDBG.Columns(iIndex).Width = (Me.ScaleWidth - dAnchoTotal) - CONTROL_MARGIN
End Sub

