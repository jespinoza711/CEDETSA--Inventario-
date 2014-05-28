VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Módulo de Inventarios"
   ClientHeight    =   8340
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8730
   Icon            =   "frmMain2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10135.03
   ScaleMode       =   0  'User
   ScaleWidth      =   8991.296
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8085
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Usuario"
            TextSave        =   "Usuario"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Base de Datos"
            TextSave        =   "Base de Datos"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Servidor"
            TextSave        =   "Servidor"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   480
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      BackColor       =   15917501
      MenuAnimations  =   3
      ToolBarsCount   =   3
      ToolsCount      =   23
      LargeIcons      =   -1  'True
      Tools           =   "frmMain2.frx":030A
      ToolBars        =   "frmMain2.frx":12778
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8040
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCaptacion_Click()

End Sub

Private Sub cmdproyprod_Click()
End Sub

Private Sub cmdproySuc_Click()

End Sub




Private Sub Command1_Click()
frmRecibeDatosComision.Show vbModal
End Sub



Private Sub Command2_Click()
    frmCalculoComision.Show vbModal
End Sub

Private Sub Command3_Click()
    frmComisiones.Show vbModal
End Sub

Private Sub Form_Load()
'lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

StatusBar1.Panels(2).Text = gsUSUARIO
StatusBar1.Panels(4).Text = gsNombreBaseDatos
StatusBar1.Panels(6).Text = gsNombreServidor
gTasaCambio = 21.5
'  lbok = CargaParametros()
'  If Not lbok Then
'
'    lbok = Mensaje("El módulo no está configurado, por favor llame a Sistemas", ICO_ERROR, False)
'    End
'
'  End If
    'Marchoso1.Filename = App.Path & "\zComisionanimada.gif"
End Sub



Private Sub Form_Unload(Cancel As Integer)
Destructor
End Sub



Private Sub SSActiveToolBars1_ToolClick(ByVal tool As ActiveToolBars.SSTool)
Dim lbok  As Boolean
Dim frm As New frmCatalogos



    Select Case tool.ID

        Case "ID_CATALOGOS"
            If UserMayAccess(gsUSUARIO, A_CATALOGOS, C_MODULO) Then

                frmCatalogos.Show vbModal

            Else
                lbok = Mensaje("Ud no tiene privilegios para entrar a esta opcion", ICO_ADVERTENCIA, False)
            End If
        Case "ID_PRODUCTOS"
'            If UserMayAccess(gsUSUARIO, A_ID_STATUS, C_MODULO) Then
'
'
                frmProductos.Show vbModal
                
'
'
'            Else
'                lbok = Mensaje("Ud no tiene privilegios para entrar a esta opcion", ICO_ADVERTENCIA, False)
'            End If
''
        Case "ID_SUPERVISORES"
'            If UserMayAccess(gsUSUARIO, A_ID_SUPERVISOR, C_MODULO) Then
'                frm.gsCatalogoName = "SUPERVISOR"
            Dim ofrmBodega As New frmBodega
            ofrmBodega.gsFormCaption = "Bodega"
            ofrmBodega.gsTitle = "Catálogo de Bodegas"
            ofrmBodega.Show
'            Else
'                lbok = Mensaje("Ud no tiene privilegios para entrar a esta opcion", ICO_ADVERTENCIA, False)
'            End If
'
        Case "ID_PREPARACION"
'            If UserMayAccess(gsUSUARIO, A_ID_CANAL, C_MODULO) Then
                Dim ofrmVendedor As New frmVendedor
                ofrmVendedor.gsFormCaption = "Vendedor"
                ofrmVendedor.gsTitle = "Maestro de Vendedores"
                ofrmVendedor.Show
'            Else
'                lbOk = Mensaje("Ud no tiene privilegios para entrar a esta opcion", ICO_ADVERTENCIA, False)
'            End If
'
        Case "ID_RESPONSABLE"
'            If UserMayAccess(gsUSUARIO, A_ID_RESPONSABLE, C_MODULO) Then
            Dim ofrmTran As New frmTransacciones
            ofrmTran.gsFormCaption = "Transacion"
            ofrmTran.gsTitle = "Transacion"
            ofrmTran.Show vbModal
'                frm.Show vbModal
'            Else
'                lbok = Mensaje("Ud no tiene privilegios para entrar a esta opcion", ICO_ADVERTENCIA, False)
'            End If
'        Case "ID_CLIENTE"
'            If UserMayAccess(gsUSUARIO, A_ID_CLIENTE, C_MODULO) Then
'                frmCliente.Show vbModal
'            Else
'                lbok = Mensaje("Ud no tiene privilegios para entrar a esta opcion", ICO_ADVERTENCIA, False)
'            End If
        Case "ID_SALIR"
                Destructor
                End
''
        Case "ID_REPORTES"
            frmTest.Show vbModal
'            If UserMayAccess(gsUSUARIO, A_ID_REPORTES, C_MODULO) Then
'                If ExistePreparacion(gsUSUARIO) Then
'                    frmReportesCC.Show vbModal
'                Else
'                    frmPrepare.Show vbModal
'                End If
'             Else
'                lbok = Mensaje("Ud no tiene privilegios para entrar a esta opcion", ICO_ADVERTENCIA, False)
'            End If
'
        Case "ID_CLASIFICACION"
            Dim ofrmMasterLotes As New frmMasterLotes
            ofrmMasterLotes.gsFormCaption = "Maestro de Lotes"
            ofrmMasterLotes.gsTitle = "Maestro de Lotes"
            ofrmMasterLotes.Show vbModal
'        Case "ID_PASSWORD"
'            If UserMayAccess(gsUSUARIO, A_PASSWORD, C_MODULO) Then
'
'
'                frmCambiaPassword.Show vbModal
'
'            Else
'                lbok = Mensaje("Ud no tiene privilegios para entrar a esta opcion", ICO_ADVERTENCIA, False)
'            End If
'
'        Case "ID_PRIVILEGIO"
'            If UserMayAccess(gsUSUARIO, A_SEGURIDAD, C_MODULO) Then
'                frmSeguridad.Show vbModal
'            Else
'                lbok = Mensaje("Ud no tiene privilegios para entrar a esta opcion", ICO_ADVERTENCIA, False)
'            End If
'
'        Case "ID_USUARIO"
'
'            If UserMayAccess(gsUSUARIO, A_USUARIOS, C_MODULO) Then
'                frmUsuario.Show vbModal
'
'            Else
'                lbok = Mensaje("Ud no tiene privilegios para entrar a esta opcion", ICO_ADVERTENCIA, False)
'            End If
'        Case "ID_ANTICIPO"
'
'            If UserMayAccess(gsUSUARIO, A_ID_ANTICIPO, C_MODULO) Then
'             frmAnticipos.Show vbModal
'            Else
'                lbok = Mensaje("Ud no tiene privilegios para entrar a esta opcion", ICO_ADVERTENCIA, False)
'            End If
'
'        Case "ID_PARAMETROS"
'
'            If UserMayAccess(gsUSUARIO, A_PARAMETROS, C_MODULO) Then
'                frmParametros.Show vbModal
'
'            Else
'                lbok = Mensaje("Ud no tiene privilegios para entrar a esta opcion", ICO_ADVERTENCIA, False)
'            End If
'        Case "ID_DIFERENCIAL"
'
'            If UserMayAccess(gsUSUARIO, A_ID_DIFERENCIAL, C_MODULO) Then
'
'                frmPrepDC.gsModulo = "CC"
'                frmPrepDC.Show vbModal
'
'            Else
'                lbok = Mensaje("Ud no tiene privilegios para entrar a esta opcion", ICO_ADVERTENCIA, False)
'            End If
'
'        Case "ID_DIFCAMBCP"
'
'            If UserMayAccess(gsUSUARIO, A_ID_DIFCAMBCP, C_MODULO) Then
'
'                frmPrepDC.gsModulo = "CP"
'                frmPrepDC.Show vbModal
'
'            Else
'                lbok = Mensaje("Ud no tiene privilegios para entrar a esta opcion", ICO_ADVERTENCIA, False)
'            End If
'
'
'
'        Case "ID_RECIBO"
'            If UserMayAccess(gsUSUARIO, A_RECIBOS, C_MODULO) Then
'                frmReciboFaltante.Show vbModal
'
'            Else
'                lbok = Mensaje("Ud no tiene privilegios para entrar a esta opcion", ICO_ADVERTENCIA, False)
'            End If
'
'        Case "ID_DIFCONSOL"
'            If UserMayAccess(gsUSUARIO, A_ID_DIFCONSOL, C_MODULO) Then
'                frmRepDifConsol.Show vbModal
'
'            Else
'                lbok = Mensaje("Ud no tiene privilegios para entrar a esta opcion", ICO_ADVERTENCIA, False)
'            End If
'
'
'        Case "ID_INTERFAZ"
'        '    If UserMayAccess(gsUSUARIO, A_ID_DIFCONSOL, C_MODULO) Then
'                frmAS400Liquidacion.Show vbModal
'
'        '    Else
'        '        lbok = Mensaje("Ud no tiene privilegios para entrar a esta opcion", ICO_ADVERTENCIA, False)
'        '    End If
'
'
'        Case "ID_FALTANTE"
'        '    If UserMayAccess(gsUSUARIO, A_ID_DIFCONSOL, C_MODULO) Then
'                frmConsultaFaltantes.Show vbModal
'
'        '    Else
'        '        lbok = Mensaje("Ud no tiene privilegios para entrar a esta opcion", ICO_ADVERTENCIA, False)
'        '    End If
'
'
'        Case "ID_CARGA"
'            frmAplicacionMasiva.Show vbModal
'
'        Case "ID_CAMBSTATUS"
'             frmCambEstadoLiq.Show vbModal
'
    End Select
Set frm = Nothing
End Sub
