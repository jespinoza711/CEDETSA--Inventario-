VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmBodega 
   BackColor       =   &H00FEE3DA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bodegas"
   ClientHeight    =   6825
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   9660
   Icon            =   "frmBodega.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FEE3DA&
      Height          =   765
      Left            =   30
      TabIndex        =   9
      Top             =   390
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
         Picture         =   "frmBodega.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Modifica los datos mostrados en el Grid con los datos digitados ..."
         Top             =   150
         Width           =   555
      End
      Begin VB.CommandButton cmdSave 
         Enabled         =   0   'False
         Height          =   555
         Left            =   1860
         Picture         =   "frmBodega.frx":1594
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Picture         =   "frmBodega.frx":325E
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Picture         =   "frmBodega.frx":3F28
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Picture         =   "frmBodega.frx":4BF2
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Deshacer / Cancelar"
         Top             =   150
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FEE3DA&
      Height          =   1335
      Left            =   210
      TabIndex        =   1
      Top             =   1260
      Width           =   9150
      Begin VB.CheckBox chkFactura 
         BackColor       =   &H00FEE3DA&
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
         Left            =   210
         TabIndex        =   0
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtBodega 
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
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtDescrBodega 
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
         Left            =   3930
         TabIndex        =   3
         Top             =   360
         Width           =   5115
      End
      Begin VB.CheckBox chkActivo 
         BackColor       =   &H00FEE3DA&
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
         Left            =   7770
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         TabIndex        =   6
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
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
         Left            =   2640
         TabIndex        =   5
         Top             =   360
         Width           =   1125
      End
   End
   Begin TrueOleDBGrid60.TDBGrid TDBG 
      Height          =   3915
      Left            =   210
      OleObjectBlob   =   "frmBodega.frx":58BC
      TabIndex        =   7
      Top             =   2640
      Width           =   9165
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
      Left            =   -690
      TabIndex        =   8
      Top             =   0
      Width           =   10710
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
            txtBodega.Enabled = True
            txtDescrBodega.Enabled = True
            chkActivo.Enabled = True
            chkActivo.value = 1
            chkFactura.Enabled = True
            chkFactura.value = 0
            txtBodega.Text = "100"
            txtDescrBodega.Text = ""
            fmtTextbox txtBodega, "R"
            fmtTextbox txtDescrBodega, "O"
            Me.TDBG.Enabled = False
        Case TypAccion.Edit
            txtDescrBodega.Enabled = True
            fmtTextbox txtBodega, "R"
            fmtTextbox txtDescrBodega, "O"
            chkActivo.Enabled = True
            chkFactura.Enabled = True
            Me.TDBG.Enabled = False
        Case TypAccion.View
            fmtTextbox txtDescrBodega, "R"
            fmtTextbox txtBodega, "R"
            fmtTextbox txtDescrBodega, "R"
            chkActivo.Enabled = False
            chkFactura.Enabled = False
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
Else
    txtBodega.Text = ""
    txtDescrBodega.Text = ""
End If

End Sub

Private Sub cmdEliminar_Click()
    Dim lbOk As Boolean
    Dim sMsg As String
    Dim sTipo As String
    Dim sFiltro As String
    Dim sActivo As String
    Dim sFactura As String

    If txtBodega.Text = "" Then
        lbOk = Mensaje("La Bodega no puede estar en Blanco", ICO_ERROR, False)
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
    lbOk = Mensaje("Está seguro de eliminar la Bodega " & rst("IDBodega").value, ICO_PREGUNTA, True)
    If lbOk Then
                lbOk = invUpdateBodega("D", txtBodega.Text, txtDescrBodega.Text, sActivo, sFactura, True)
        
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
    If txtBodega.Text = "" Then
        lbOk = Mensaje("La Bodega no puede estar en Blanco", ICO_ERROR, False)
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
        lbOk = Mensaje("La Descripción del Centro no puede estar en blanco", ICO_ERROR, False)
        Exit Sub
    End If
    

        
    If Accion = Add Then
    
        If Not (rst.EOF And rst.BOF) Then
            sFiltro = "IDBodega = '" & txtBodega.Text & "'"
            If ExiteRstKey(rst, sFiltro) Then
               lbOk = Mensaje("Ya exista Bodega ", ICO_ERROR, False)
                txtBodega.SetFocus
            Exit Sub
            End If
        End If
    
            lbOk = invUpdateBodega("I", txtBodega.Text, txtDescrBodega.Text, sActivo, sFactura, True)
            
            If lbOk Then
                sMsg = "La Bodega ha sido registrada exitosamente ... "
                lbOk = Mensaje(sMsg, ICO_OK, False)
                ' actualiza datos
                cargaGrid
                Accion = View
                HabilitarControles
                HabilitarBotones
            Else
                 sMsg = "Ha ocurrido un error tratando de Agregar la Bodega... "
                lbOk = Mensaje(sMsg, ICO_ERROR, False)
            End If
    End If ' si estoy adicionando
        If Accion = Edit Then
            If Not (rst.EOF And rst.BOF) Then
                lbOk = invUpdateBodega("U", txtBodega.Text, txtDescrBodega.Text, sActivo, sFactura, True)
                If lbOk Then
                    sMsg = "La Bodega ha sido registrada exitosamente ... "
                    lbOk = Mensaje(sMsg, ICO_OK, False)
                    ' actualiza datos
                    cargaGrid
                    Accion = View
                    HabilitarControles
                    HabilitarBotones
                Else
                    sMsg = "Ha ocurrido un error tratando de Agregar la Bodega... "
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
