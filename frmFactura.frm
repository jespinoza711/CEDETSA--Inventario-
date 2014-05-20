VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFactura 
   Caption         =   "Form1"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15480
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   15480
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   180
      TabIndex        =   35
      Top             =   1410
      Width           =   13935
      Begin VB.CommandButton cmdCliente 
         Height          =   320
         Left            =   960
         Picture         =   "frmFactura.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   240
         Width           =   300
      End
      Begin VB.TextBox txtNombres 
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
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   240
         Width           =   6855
      End
      Begin VB.TextBox txtCodCliente 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   38
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optCredito 
         Caption         =   "Crédito"
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
         Left            =   10920
         TabIndex        =   37
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optContado 
         Caption         =   "Contado"
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
         Left            =   12240
         TabIndex        =   36
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente :"
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
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Factura :"
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
         Left            =   9480
         TabIndex        =   41
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdCobro 
      Height          =   495
      Left            =   14700
      Picture         =   "frmFactura.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3210
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdDelItem 
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
      Height          =   495
      Left            =   14700
      Picture         =   "frmFactura.frx":0784
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Elimina el item actualmente seleccionado en el grid de datos ..."
      Top             =   4050
      Width           =   495
   End
   Begin VB.CommandButton cmdSave 
      Enabled         =   0   'False
      Height          =   495
      Left            =   14700
      Picture         =   "frmFactura.frx":0BC6
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Aplica y Guarda los datos de la transacción en Firme ..."
      Top             =   2490
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdAddItem 
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
      Height          =   495
      Left            =   14700
      Picture         =   "frmFactura.frx":0ED0
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Inicializa los controles para Agregar otro item ..."
      Top             =   6030
      Width           =   495
   End
   Begin VB.CommandButton cmdEditItem 
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
      Height          =   495
      Left            =   14700
      Picture         =   "frmFactura.frx":11DA
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Modifica los datos mostrados en el Grid con los datos digitados ..."
      Top             =   5190
      Width           =   495
   End
   Begin VB.CommandButton cmdOkCO 
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
      Height          =   495
      Left            =   14700
      Picture         =   "frmFactura.frx":1AA4
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Aprueba, Aplica los datos digitados para el item en proceso y son ingresados en el grid de datos ..."
      Top             =   6810
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   60
      TabIndex        =   15
      Top             =   90
      Width           =   13935
      Begin VB.TextBox txtCodCajero 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   22
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtDescrCajero 
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
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   840
         Width           =   4095
      End
      Begin VB.CommandButton cmdCajero 
         Height          =   320
         Left            =   1320
         Picture         =   "frmFactura.frx":1DAE
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   840
         Width           =   300
      End
      Begin VB.CommandButton cmdBodega 
         Height          =   320
         Left            =   1320
         Picture         =   "frmFactura.frx":20F0
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   300
      End
      Begin VB.TextBox txtDescrBodega 
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
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   360
         Width           =   4095
      End
      Begin VB.TextBox txtcodBodega 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   17
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdPedidos 
         Caption         =   "Cargar Pedidos"
         Height          =   375
         Left            =   12240
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPFecFac 
         Height          =   255
         Left            =   9600
         TabIndex        =   23
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   """DD/MM/YYYY HH:MM"""
         Format          =   61472769
         CurrentDate     =   38090.4465277778
      End
      Begin VB.Label lblVendedor 
         Caption         =   "Vendedor :"
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
         Left            =   120
         TabIndex        =   28
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Bodega :"
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
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha Factura :"
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
         Left            =   7320
         TabIndex        =   26
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblTipoCambioFac 
         Alignment       =   1  'Right Justify
         Caption         =   "25.3525"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9600
         TabIndex        =   25
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "Tipo de Cambio Factura :"
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
         Left            =   7320
         TabIndex        =   24
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   300
      TabIndex        =   0
      Top             =   5490
      Width           =   8895
      Begin VB.TextBox txtPrecioDcto 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtDescuento 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4080
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtImpuesto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         TabIndex        =   5
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtCodProd 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtDescProd 
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
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   5895
      End
      Begin VB.CommandButton cmdProducto 
         Height          =   320
         Left            =   960
         Picture         =   "frmFactura.frx":2432
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   300
      End
      Begin VB.Label Label10 
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3000
         TabIndex        =   14
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Descuento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4080
         TabIndex        =   13
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Impuesto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5400
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.Label txtTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   6600
         TabIndex        =   11
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Precio Vta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6600
         TabIndex        =   10
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad"
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
         TabIndex        =   9
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblProducto 
         Caption         =   "Producto:"
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
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
   End
   Begin TrueOleDBGrid60.TDBGrid TDBGFAC 
      Height          =   2895
      Left            =   420
      OleObjectBlob   =   "frmFactura.frx":2774
      TabIndex        =   43
      Top             =   2250
      Width           =   13695
   End
   Begin MSComctlLib.StatusBar StatusSubTotal 
      Height          =   495
      Left            =   9540
      TabIndex        =   44
      Top             =   5490
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Sub Total"
            TextSave        =   "Sub Total"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusTotal 
      Height          =   495
      Left            =   9540
      TabIndex        =   45
      Top             =   7530
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Total"
            TextSave        =   "Total"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusDescuento 
      Height          =   495
      Left            =   9540
      TabIndex        =   46
      Top             =   6090
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Descuento"
            TextSave        =   "Descuento"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusImpuesto 
      Height          =   495
      Left            =   9540
      TabIndex        =   47
      Top             =   6810
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Impuesto"
            TextSave        =   "Impuesto"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgEdit 
      Height          =   480
      Left            =   14220
      Picture         =   "frmFactura.frx":7270
      Top             =   5250
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgAdd 
      Height          =   480
      Left            =   14220
      Picture         =   "frmFactura.frx":76B2
      Top             =   6090
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgOk 
      Height          =   480
      Left            =   14220
      Picture         =   "frmFactura.frx":7AF4
      Top             =   6810
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

