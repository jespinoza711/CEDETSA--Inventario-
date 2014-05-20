VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmEscalaBonificacion 
   Caption         =   "Form1"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   10005
   StartUpPosition =   3  'Windows Default
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
      Left            =   8730
      Picture         =   "frmEscalaBonificacion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Modifica los datos mostrados en el Grid con los datos digitados ..."
      Top             =   3240
      Width           =   555
   End
   Begin VB.CommandButton cmdSave 
      Enabled         =   0   'False
      Height          =   555
      Left            =   8730
      Picture         =   "frmEscalaBonificacion.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Aplica y Guarda los datos de la transacción en Firme ..."
      Top             =   4440
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
      Left            =   8730
      Picture         =   "frmEscalaBonificacion.frx":2994
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Agrega el item con los datos digitados..."
      Top             =   2640
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
      Left            =   8730
      Picture         =   "frmEscalaBonificacion.frx":365E
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Elimina el item actualmente seleccionado en el grid de datos ..."
      Top             =   3840
      Width           =   555
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   690
      TabIndex        =   1
      Top             =   975
      Width           =   8775
      Begin VB.TextBox txtProducto 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002F2F2F&
         Height          =   315
         Left            =   1440
         TabIndex        =   10
         Top             =   345
         Width           =   1050
      End
      Begin VB.TextBox txtDescr 
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
         Height          =   315
         Left            =   2565
         TabIndex        =   9
         Top             =   345
         Width           =   6030
      End
      Begin VB.Frame Frame4 
         Caption         =   "Bonificación"
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   8415
         Begin VB.TextBox txtPorCada 
            BackColor       =   &H8000000F&
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
            Left            =   3960
            TabIndex        =   5
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtBonifica 
            BackColor       =   &H8000000F&
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
            Left            =   6720
            TabIndex        =   4
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtIDEscala 
            BackColor       =   &H8000000F&
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
            TabIndex        =   3
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label12 
            Caption         =   "Por Cada :"
            ForeColor       =   &H002F2F2F&
            Height          =   255
            Left            =   2640
            TabIndex        =   8
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "Bonifica :"
            ForeColor       =   &H002F2F2F&
            Height          =   255
            Left            =   5520
            TabIndex        =   7
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Escala :"
            ForeColor       =   &H002F2F2F&
            Height          =   255
            Left            =   480
            TabIndex        =   6
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Producto :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002F2F2F&
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
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
      Left            =   8730
      Picture         =   "frmEscalaBonificacion.frx":4328
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Deshacer / Cancelar"
      Top             =   5040
      Width           =   555
   End
   Begin TrueOleDBGrid60.TDBGrid TDBG 
      Height          =   2955
      Left            =   660
      OleObjectBlob   =   "frmEscalaBonificacion.frx":4FF2
      TabIndex        =   16
      Top             =   2640
      Width           =   7905
   End
   Begin VB.Label lbFormCaption 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Left            =   0
      TabIndex        =   17
      Top             =   240
      Width           =   10140
   End
   Begin VB.Image Image1 
      Height          =   765
      Left            =   -270
      Picture         =   "frmEscalaBonificacion.frx":AB6B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11490
   End
End
Attribute VB_Name = "frmEscalaBonificacion"
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
Public gsPorCada As String
Public gsBonifica As String
Public gsFormCaption As String
Public gsTitle As String
Public gbOnlyShow As Boolean
Public giCantidadFuente As Integer
Public gsIDProducto As String
Public gsDescr As String

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
End Sub

Public Sub HabilitarControles()
    Select Case Accion
        Case TypAccion.Add
            txtProducto.Enabled = True
            txtDescr.Enabled = True
            txtPorCada.Enabled = True
            txtPorCada.Enabled = True
            txtIDEscala.Text = "1"
            txtIDEscala.Enabled = False
            txtPorCada.Text = ""
            txtBonifica.Text = ""
            fmtTextbox txtProducto, "R"
            fmtTextbox txtDescr, "R"
            fmtTextbox txtPorCada, "O"
            fmtTextbox txtBonifica, "O"
            TDBG.Enabled = False
        Case TypAccion.Edit
            txtProducto.Enabled = True
            txtDescr.Enabled = True
            fmtTextbox txtProducto, "R"
            fmtTextbox txtDescr, "R"
            txtPorCada.Enabled = True
            txtBonifica.Enabled = True
            txtIDEscala.Enabled = False
            fmtTextbox txtPorCada, "O"
            fmtTextbox txtBonifica, "O"
            TDBG.Enabled = False
        Case TypAccion.View
            fmtTextbox txtProducto, "R"
            fmtTextbox txtDescr, "R"
            fmtTextbox txtIDEscala, "R"
            fmtTextbox txtPorCada, "R"
            fmtTextbox txtBonifica, "R"
            TDBG.Enabled = True
    End Select
End Sub


Private Sub cmdAdd_Click()
    Accion = Add
    HabilitarBotones
    HabilitarControles
    txtPorCada.SetFocus
End Sub

Private Sub cmdEditItem_Click()
    Accion = Edit
    GetDataFromGridToControl
    HabilitarBotones
    HabilitarControles
End Sub
Private Sub GetDataFromGridToControl()
    If Not (rst.EOF And rst.BOF) Then
        txtProducto.Text = rst("IDProducto").value
        txtDescr.Text = rst("Descr").value
        txtIDEscala.Text = rst("IDEscala").value
        txtPorCada.Text = rst("PorCada").value
        txtBonifica.Text = rst("Bonifica").value
    Else
        txtPorCada.Text = "0"
        txtBonifica.Text = "0"
    End If
End Sub

Private Sub cmdEliminar_Click()
    Dim lbok As Boolean
    Dim sMsg As String
    Dim sTipo As String
    Dim sFiltro As String
    Dim sActivo As String
    Dim sFactura As String
    
        If txtProducto.Text = "" Then
            lbok = Mensaje("El Vendedor no puede estar en Blanco", ICO_ERROR, False)
            Exit Sub
        End If

        
        If txtPorCada.Text = "" Then
            lbok = Mensaje("El valor Por Cada no puede estar en Blanco", ICO_ERROR, False)
            Exit Sub
        End If
        
        If txtBonifica.Text = "" Then
            lbok = Mensaje("El valor Bonifica no puede estar en Blanco", ICO_ERROR, False)
            Exit Sub
        End If
        
        If txtIDEscala.Text = "" Then
            lbok = Mensaje("El valor Escala no puede estar en Blanco", ICO_ERROR, False)
            Exit Sub
        End If
        
        ' hay que validar la integridad referencial
        lbok = Mensaje("Está seguro de eliminar La Escala " & rst("IDEscala").value, ICO_PREGUNTA, True)
        If lbok Then
                    lbok = fafUpdateEscalaBonificacion("D", txtProducto.Text, txtIDEscala.Text, txtPorCada.Text, txtBonifica.Text)
            
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
        If txtProducto.Text = "" Then
            lbok = Mensaje("El Vendedor no puede estar en Blanco", ICO_ERROR, False)
            Exit Sub
        End If

        If txtIDEscala.Text = "" Then
            lbok = Mensaje("El Valor de la Escala no puede estar en Blanco", ICO_ERROR, False)
            Exit Sub
        End If

        If txtPorCada.Text = "" Then
            lbok = Mensaje("El Valor Por Cada no puede estar en Blanco", ICO_ERROR, False)
            Exit Sub
        End If
        
        If txtBonifica.Text = "" Then
            lbok = Mensaje("El Valor Bonificación no puede estar en blanco", ICO_ERROR, False)
            Exit Sub
        End If
        
    If Not Val_TextboxNum(txtPorCada) Then
        lbok = Mensaje("El valor Por Cada debe ser Numérico", ICO_ERROR, False)
        
        Exit Sub
    End If
    
    If Not Val_TextboxNum(txtBonifica) Then
        lbok = Mensaje("El valor Bonifica debe ser Numérico", ICO_ERROR, False)
        
        Exit Sub
    End If
    
    If Val(txtBonifica.Text) > Val(txtPorCada.Text) Then
        lbok = Mensaje("El Valor Bonifica no puede ser Mayor que el Valor Por Cada", ICO_ERROR, False)
        
        Exit Sub
    End If
    
    If Val(txtBonifica.Text) = 0 Or Val(txtPorCada.Text) = 0 Then
        lbok = Mensaje("El Valor Bonifica y Por Cada no puede ser igual a Cero", ICO_ERROR, False)
        
        Exit Sub
    End If
    
    If Not EsEntero(txtPorCada.Text) Then
        lbok = Mensaje("El Valor Por Cada  tiene que ser entero", ICO_ERROR, False)
        Exit Sub
    End If
            
    If Not EsEntero(txtBonifica.Text) Then
        lbok = Mensaje("El Valor Bonifica  tiene que ser entero", ICO_ERROR, False)
        Exit Sub
    End If
            
            
    If Accion = Add Then
    
        If Not (rst.EOF And rst.BOF) Then
            sFiltro = "IDProducto = " & txtProducto.Text & " and PorCada= " & txtPorCada.Text & " and Bonifica = " & txtBonifica.Text
            If ExiteRstKey(rst, sFiltro) Then
               lbok = Mensaje("Ya existe esa escala ", ICO_ERROR, False)
                txtPorCada.SetFocus
            Exit Sub
            End If
        End If
    
        lbok = fafUpdateEscalaBonificacion("I", txtProducto.Text, txtIDEscala.Text, txtPorCada.Text, txtBonifica.Text)
        
        If lbok Then
            sMsg = "La escala ha sido registrada exitosamente ... "
            lbok = Mensaje(sMsg, ICO_OK, False)
            ' actualiza datos
            Accion = View
            cargaGrid
            HabilitarControles
            HabilitarBotones
        Else
            sMsg = "Ha ocurrido un error tratando de Agregar el Vendedor... "
            lbok = Mensaje(sMsg, ICO_ERROR, False)
        End If
      
    End If ' si estoy adicionando
    If Accion = Edit Then
        If Not (rst.EOF And rst.BOF) Then
            lbok = lbok = fafUpdateEscalaBonificacion("I", txtProducto.Text, txtIDEscala.Text, txtPorCada.Text, txtBonifica.Text)
            If lbok Then
                sMsg = "Los datos fueron grabados Exitosamente ... "
                lbok = Mensaje(sMsg, ICO_OK, False)
                ' actualiza datos
                Accion = View
                cargaGrid
                HabilitarControles
                HabilitarBotones
            Else
                sMsg = "Ha ocurrido un error tratando de Actualizar el Vendedor... "
                lbok = Mensaje(sMsg, ICO_ERROR, False)
            End If
        End If
        
    End If ' si estoy adicionando

End Sub


Private Sub cmdUndo_Click()
    GetDataFromGridToControl
    Accion = View
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
    'sFormCaption = "Catalogo de Vendedores"
    Caption = gsFormCaption
    lbFormCaption = gsTitle
    Accion = View
    HabilitarBotones
    HabilitarControles
    cargaGrid
    txtProducto.Text = gsIDProducto
    txtDescr.Text = gsDescr
    If gbOnlyShow = True Then
        cmdEditItem.Enabled = False
        cmdSave.Enabled = False
        cmdUndo.Enabled = False
        cmdEliminar.Enabled = False
        cmdAdd.Enabled = False
        lbFormCaption.Caption = "Seleccione un elemento dando doble click"
    End If
    
End Sub


Private Sub cargaGrid()
    Dim sIndependiente As String
    If rst.State = adStateOpen Then rst.Close
    rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
    rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
    rst.CursorLocation = adUseClient ' Cursor local al cliente
    rst.LockType = adLockOptimistic
    GSSQL = gsCompania & ".fafgetEscalaBonificacion " & gsIDProducto
    If rst.State = adStateOpen Then rst.Close
    Set rst = GetRecordset(GSSQL)
    If Not (rst.EOF And rst.BOF) Then
      Set TDBG.DataSource = rst
      TDBG.Refresh
    End If
End Sub


Private Sub TDBG_DblClick()
Dim lbok As Boolean
If Not (rst.EOF And rst.BOF) Then
    If gbOnlyShow = True Then
        If giCantidadFuente > rst("PorCada").value Then
            lbok = Mensaje("Ud ha seleccionado una Escala que no suple la Cantidad requerida. Por favor seleccione otra", ICO_ERROR, False)
            gsPorCada = "0"
            gsBonifica = "0"
            Exit Sub
        End If
    
        gsPorCada = rst("PorCada").value
        gsBonifica = rst("Bonifica").value
        Hide
    End If
End If
End Sub

Private Sub TDBG_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    GetDataFromGridToControl
    If Not gbOnlyShow Then
        HabilitarBotones
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (rst Is Nothing) Then Set rst = Nothing
End Sub

