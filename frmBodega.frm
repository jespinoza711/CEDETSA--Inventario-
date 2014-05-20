VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmBodega 
   Caption         =   "Bodegas"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10980
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBodega.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   10980
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdUsuario 
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
      Left            =   8100
      Picture         =   "frmBodega.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Agrega el item con los datos digitados..."
      Top             =   2190
      Width           =   555
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   210
      TabIndex        =   10
      Top             =   1170
      Width           =   7695
      Begin VB.CheckBox chkActivo 
         Caption         =   "Activo ?"
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
         Left            =   5160
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtDescrBodega 
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
         Left            =   3360
         TabIndex        =   20
         Top             =   360
         Width           =   4095
      End
      Begin VB.TextBox txtBodega 
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
         Left            =   1200
         TabIndex        =   19
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox chkFactura 
         Caption         =   "Se Factura en esta Bodega ?"
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
         Left            =   2040
         TabIndex        =   18
         Top             =   720
         Width           =   2895
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ultimos Consecutivos"
         Height          =   615
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   7335
         Begin VB.TextBox txtConsecPedido 
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
            Left            =   1680
            TabIndex        =   14
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtConsecFactura 
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
            Left            =   3600
            TabIndex        =   13
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtPrefFactura 
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
            Left            =   6240
            TabIndex        =   12
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Pedido Factura:"
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
            TabIndex        =   17
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label3 
            Caption         =   " Factura:"
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
            Left            =   2760
            TabIndex        =   16
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Label5 
            Caption         =   " Prefijo Factura:"
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
            Left            =   4920
            TabIndex        =   15
            Top             =   240
            Width           =   1275
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Descripción:"
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
         Left            =   2160
         TabIndex        =   23
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Bodega:"
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
         TabIndex        =   22
         Top             =   360
         Width           =   795
      End
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   10980
      TabIndex        =   6
      Top             =   0
      Width           =   10980
      Begin VB.Image Image 
         Height          =   645
         Index           =   2
         Left            =   60
         Picture         =   "frmBodega.frx":0BD4
         Stretch         =   -1  'True
         Top             =   60
         Width           =   720
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Catálogo de Bodegas"
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
         TabIndex        =   8
         Top             =   420
         Width           =   1320
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
         TabIndex        =   7
         Top             =   90
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   90
      TabIndex        =   0
      Top             =   270
      Visible         =   0   'False
      Width           =   9585
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
         Picture         =   "frmBodega.frx":1C8B
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Modifica los datos mostrados en el Grid con los datos digitados ..."
         Top             =   150
         Width           =   555
      End
      Begin VB.CommandButton cmdSave 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1860
         Picture         =   "frmBodega.frx":2955
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Aplica y Guarda los datos de la transacción en Firme ..."
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
         Picture         =   "frmBodega.frx":461F
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Agrega el item con los datos digitados..."
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
         Picture         =   "frmBodega.frx":52E9
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Elimina el item actualmente seleccionado en el grid de datos ..."
         Top             =   150
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
         Left            =   2460
         Picture         =   "frmBodega.frx":5FB3
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Deshacer / Cancelar"
         Top             =   150
         Width           =   555
      End
   End
   Begin Inventario.CtlLiner CtlLiner 
      Height          =   30
      Left            =   0
      TabIndex        =   9
      Top             =   750
      Width           =   17925
      _ExtentX        =   31618
      _ExtentY        =   53
   End
   Begin TrueOleDBGrid60.TDBGrid TDBG 
      Height          =   4335
      Left            =   180
      OleObjectBlob   =   "frmBodega.frx":6C7D
      TabIndex        =   24
      Top             =   2910
      Width           =   8385
   End
End
Attribute VB_Name = "frmBodega"
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
            If rst.State = adStateClosed Then
                cmdEditItem.Enabled = False
                cmdSave.Enabled = False
                cmdUndo.Enabled = False
                cmdEliminar.Enabled = False
                cmdAdd.Enabled = True
                Exit Sub
            End If
            If rst.RecordCount <> 0 Then
                cmdSave.Enabled = False
                cmdUndo.Enabled = False
                cmdEliminar.Enabled = True
                cmdAdd.Enabled = True
                cmdEditItem.Enabled = True
            Else
                cmdAdd.Enabled = True
                cmdEditItem.Enabled = False
                cmdSave.Enabled = False
                cmdUndo.Enabled = False
                cmdEliminar.Enabled = False
                cmdAdd.Enabled = True
            End If
    End Select
    HabilitarBotonesMain
End Sub

Public Sub HabilitarControles()
    Select Case Accion
         Case TypAccion.Add
            txtBodega.Enabled = True
            txtDescrBodega.Enabled = True
            txtPrefFactura.Enabled = True
            chkActivo.Enabled = True
            chkActivo.value = 1
            chkFactura.Enabled = True
            chkFactura.value = 0
            txtBodega.Text = "100"
            txtDescrBodega.Text = ""
            fmtTextbox txtBodega, "R"
            fmtTextbox txtDescrBodega, "O"
            txtConsecPedido.Text = "0"
            fmtTextbox txtConsecPedido, "R"
            txtConsecFactura.Text = "0"
            fmtTextbox txtConsecFactura, "R"
            txtPrefFactura.Text = ""
            fmtTextbox txtPrefFactura, "O"
            Me.TDBG.Enabled = False
        Case TypAccion.Edit
            txtDescrBodega.Enabled = True
            fmtTextbox txtBodega, "R"
            fmtTextbox txtDescrBodega, "O"
            chkActivo.Enabled = True
            chkFactura.Enabled = True
            TDBG.Enabled = False
            txtConsecFactura.Enabled = False
            txtConsecPedido.Enabled = False
            txtPrefFactura.Enabled = False
        Case TypAccion.View
            fmtTextbox txtDescrBodega, "R"
            fmtTextbox txtBodega, "R"
            fmtTextbox txtDescrBodega, "R"
            chkActivo.Enabled = False
            chkFactura.Enabled = False
            txtConsecFactura.Enabled = False
            txtConsecPedido.Enabled = False
            txtPrefFactura.Enabled = False
            
            Me.TDBG.Enabled = True
    End Select
End Sub

Private Sub cmdAdd_Click()
    Accion = Add
    HabilitarBotones
    HabilitarControles
    txtDescrBodega.SetFocus
End Sub

Private Sub cmdEditItem_Click()
    Accion = Edit
    GetDataFromGridToControl
    HabilitarBotones
    HabilitarControles
End Sub
Private Sub GetDataFromGridToControl()
If Not (rst.EOF And rst.BOF) Then
    txtBodega.Text = rst("IDBodega").value
    txtDescrBodega.Text = rst("DescrBodega").value
    If rst("Activo").value = True Then
        chkActivo.value = 1
    Else
        chkActivo.value = 0
    End If
    If rst("Factura").value = True Then
        chkFactura.value = 1
    Else
        chkFactura.value = 0
    End If
    txtConsecFactura.Text = rst("ConsecFactura").value
    txtConsecPedido.Text = rst("ConsecPedido").value
    txtPrefFactura.Text = rst("PreFactura").value
    
Else
    txtBodega.Text = ""
    txtDescrBodega.Text = ""
    txtPrefFactura.Text = ""
    txtConsecFactura.Text = "0"
    txtConsecPedido.Text = "0"
End If

End Sub

Private Sub cmdEliminar_Click()
    Dim lbok As Boolean
        Dim sMsg As String
    Dim sTipo As String
    Dim sFiltro As String
    Dim sActivo As String
    Dim sFactura As String

    If txtBodega.Text = "" Then
        lbok = Mensaje("La Bodega no puede estar en Blanco", ICO_ERROR, False)
        Exit Sub
    End If
    If chkActivo.value = 1 Then
        sActivo = "1"
    Else
        sActivo = "0"

    End If
    
    If chkFactura.value = 1 Then
        sFactura = "1"
    Else
        sFactura = "0"
    End If
    
    ' hay que validar la integridad referencial
    lbok = Mensaje("Está seguro de eliminar la Bodega " & rst("IDBodega").value, ICO_PREGUNTA, True)
    If lbok Then
                lbok = invUpdateBodega("D", txtBodega.Text, txtDescrBodega.Text, sActivo, sFactura, txtPrefFactura.Text)
        
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
    If txtBodega.Text = "" Then
        lbok = Mensaje("La Bodega no puede estar en Blanco", ICO_ERROR, False)
        Exit Sub
    End If
    If chkActivo.value = 1 Then
        sActivo = "1"
    Else
        sActivo = "0"
    End If
    If chkFactura.value = 1 Then
        sFactura = "1"
    Else
        sFactura = "0"
    End If
    If txtDescrBodega.Text = "" Then
        lbok = Mensaje("La Descripción del Centro no puede estar en blanco", ICO_ERROR, False)
        Exit Sub
    End If
    If txtConsecFactura.Text = "" Then
        txtConsecFactura.Text = "0"
    End If
    If txtConsecPedido.Text = "" Then
        txtConsecPedido.Text = "0"
    End If
    If txtPrefFactura.Text = "" Then
        txtPrefFactura.Text = ""
    End If
    

        
    If Accion = Add Then
    
        If Not (rst.EOF And rst.BOF) Then
            sFiltro = "IDBodega = '" & txtBodega.Text & "'"
            If ExiteRstKey(rst, sFiltro) Then
               lbok = Mensaje("Ya exista Bodega ", ICO_ERROR, False)
                txtBodega.SetFocus
            Exit Sub
            End If
        End If
    
            lbok = invUpdateBodega("I", txtBodega.Text, txtDescrBodega.Text, sActivo, sFactura, txtPrefFactura.Text)
            
            If lbok Then
                sMsg = "La Bodega ha sido registrada exitosamente ... "
                lbok = Mensaje(sMsg, ICO_OK, False)
                ' actualiza datos
                cargaGrid
                Accion = View
                HabilitarControles
                HabilitarBotones
            Else
                 sMsg = "Ha ocurrido un error tratando de Agregar la Bodega... "
                lbok = Mensaje(sMsg, ICO_ERROR, False)
            End If
    End If ' si estoy adicionando
        If Accion = Edit Then
            If Not (rst.EOF And rst.BOF) Then
                lbok = invUpdateBodega("U", txtBodega.Text, txtDescrBodega.Text, sActivo, sFactura, txtPrefFactura.Text)
                If lbok Then
                    sMsg = "Registro actualizado exitosamente... "
                    lbok = Mensaje(sMsg, ICO_OK, False)
                    ' actualiza datos
                    cargaGrid
                    Accion = View
                    HabilitarControles
                    HabilitarBotones
                    Else
                        sMsg = "Ha ocurrido un error tratando de Actualizar la Bodega... "
                        lbok = Mensaje(sMsg, ICO_ERROR, False)
                    
                End If
            End If
        
    End If ' si estoy adicionando
End Sub

Private Sub cmdUsuario_Click()
If Not (rst.EOF And rst.BOF) Then
    Dim frm As frmBodegaUsuario
    Set frm = New frmBodegaUsuario
    frm.gsDescrBodega = txtDescrBodega.Text
    frm.gsIDBodega = txtBodega.Text
    frm.gsFormCaption = "Usuarios con Acceso a Bodegas"
    frm.gsTitle = "Asignación de Usuarios a Bodegas"
    frm.Show vbModal
    Set frm = Nothing
End If
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



Private Sub cargaGrid()
    Dim sIndependiente As String
    If rst.State = adStateOpen Then rst.Close
    rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
    rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
    rst.CursorLocation = adUseClient ' Cursor local al cliente
    rst.LockType = adLockOptimistic
    GSSQL = gsCompania & ".invGetBodegas -1"
    If rst.State = adStateOpen Then rst.Close
    Set rst = GetRecordset(GSSQL)
    If Not (rst.EOF And rst.BOF) Then
      Set TDBG.DataSource = rst
      'CargarDatos rst, TDBG, "Codigo", "Descr"
      TDBG.Refresh
      'IniciaIconos
    End If
End Sub


Private Sub Form_Resize()
 On Error Resume Next
    If WindowState <> vbMinimized Then
        If Me.Width < 9195 Then Me.Width = 9195
        If Me.Height < 4500 Then Me.Height = 4500
        
        center_obj_horizontal Me, Frame2
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
