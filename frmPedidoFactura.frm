VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPedidoFactura 
   Caption         =   "Form1"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16275
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   16275
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   60
      Top             =   1260
      Width           =   15015
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
         TabIndex        =   66
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtNombre 
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
         TabIndex        =   65
         Top             =   240
         Width           =   6855
      End
      Begin VB.CommandButton cmdCliente 
         Height          =   320
         Left            =   960
         Picture         =   "frmPedidoFactura.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   240
         Width           =   300
      End
      Begin VB.CommandButton cmdTipo 
         Height          =   320
         Left            =   10800
         Picture         =   "frmPedidoFactura.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   240
         Width           =   300
      End
      Begin VB.TextBox txtTipo 
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
         Left            =   11160
         TabIndex        =   62
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtDescrTipo 
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
         Left            =   12000
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   240
         Width           =   2655
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
         TabIndex        =   68
         Top             =   240
         Width           =   1215
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
         TabIndex        =   67
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCobro 
      Height          =   495
      Left            =   15360
      Picture         =   "frmPedidoFactura.frx":0684
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   3060
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
      Left            =   15360
      Picture         =   "frmPedidoFactura.frx":0AC6
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Elimina el item actualmente seleccionado en el grid de datos ..."
      Top             =   3900
      Width           =   495
   End
   Begin VB.CommandButton cmdSave 
      Enabled         =   0   'False
      Height          =   495
      Left            =   15360
      Picture         =   "frmPedidoFactura.frx":0F08
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "Aplica y Guarda los datos de la transacción en Firme ..."
      Top             =   2340
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
      Left            =   15360
      Picture         =   "frmPedidoFactura.frx":1212
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   "Inicializa los controles para Agregar otro item ..."
      Top             =   5880
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
      Left            =   15360
      Picture         =   "frmPedidoFactura.frx":151C
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "Modifica los datos mostrados en el Grid con los datos digitados ..."
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton cmdOkCO 
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
      Left            =   15360
      Picture         =   "frmPedidoFactura.frx":1DE6
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Aprueba, Aplica los datos digitados para el item en proceso y son ingresados en el grid de datos ..."
      Top             =   6660
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   720
      TabIndex        =   41
      Top             =   60
      Width           =   15135
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
         Left            =   1560
         TabIndex        =   47
         Top             =   360
         Width           =   735
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   360
         Width           =   4335
      End
      Begin VB.CommandButton cmdBodega 
         Height          =   320
         Left            =   1080
         Picture         =   "frmPedidoFactura.frx":20F0
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   360
         Width           =   300
      End
      Begin VB.CommandButton cmdVendedor 
         Height          =   320
         Left            =   1080
         Picture         =   "frmPedidoFactura.frx":2432
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   840
         Width           =   300
      End
      Begin VB.TextBox txtDescrVendedor 
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   840
         Width           =   4335
      End
      Begin VB.TextBox txtCodVendedor 
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
         Left            =   1560
         TabIndex        =   42
         Top             =   840
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DTPFecFac 
         Height          =   375
         Left            =   9600
         TabIndex        =   48
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   """DD/MM/YYYY"""
         Format          =   61603841
         CurrentDate     =   38090.4465277778
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha Pedido :"
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
         TabIndex        =   53
         Top             =   360
         Width           =   1455
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
         TabIndex        =   52
         Top             =   360
         Width           =   855
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
         TabIndex        =   51
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblPedido 
         Caption         =   "Pedido No :"
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
         Left            =   11640
         TabIndex        =   50
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblNoPedido 
         Alignment       =   1  'Right Justify
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
         Left            =   13200
         TabIndex        =   49
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   5100
      Width           =   9975
      Begin VB.CommandButton cmdProducto 
         Height          =   320
         Left            =   1080
         Picture         =   "frmPedidoFactura.frx":2774
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   200
         Width           =   300
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   240
         Width           =   7215
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
         Left            =   1440
         TabIndex        =   26
         Top             =   240
         Width           =   975
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
         Left            =   5880
         TabIndex        =   25
         Text            =   "0.0"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtPrecio 
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
         Left            =   1080
         TabIndex        =   24
         Text            =   "0"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtPorcImpuesto 
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
         Left            =   3840
         TabIndex        =   23
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtSubTotal 
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
         Left            =   1080
         TabIndex        =   22
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox txtTotalImpuesto 
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
         Left            =   5880
         TabIndex        =   21
         Text            =   "0.0"
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtTotal 
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
         Left            =   7920
         TabIndex        =   20
         Text            =   "0.0"
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox txtIDLote 
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
         Left            =   4320
         TabIndex        =   19
         Text            =   "0"
         Top             =   2040
         Width           =   855
      End
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   9495
         Begin VB.TextBox txtFABonifica 
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
            Left            =   8400
            TabIndex        =   15
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtFAPorCada 
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
            Left            =   6600
            TabIndex        =   14
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chkAplicaBonificacion 
            Caption         =   "Bonifica en esta Factura"
            Height          =   255
            Left            =   2400
            TabIndex        =   13
            Top             =   240
            Width           =   2175
         End
         Begin VB.CommandButton cmdEscala 
            Caption         =   "Escala"
            Height          =   315
            Left            =   4680
            TabIndex        =   12
            Top             =   210
            Width           =   975
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
            Left            =   1200
            TabIndex        =   11
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label11 
            Caption         =   "Bonifica :"
            ForeColor       =   &H002F2F2F&
            Height          =   255
            Left            =   7680
            TabIndex        =   18
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label12 
            Caption         =   "Por Cada :"
            ForeColor       =   &H002F2F2F&
            Height          =   255
            Left            =   5760
            TabIndex        =   17
            Top             =   240
            Width           =   735
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
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdLote 
         Height          =   320
         Left            =   3960
         Picture         =   "frmPedidoFactura.frx":2AB6
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2040
         Width           =   300
      End
      Begin VB.TextBox TotalDescuento 
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
         Height          =   330
         Left            =   3360
         TabIndex        =   8
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CheckBox chkLoteAutomaticos 
         Caption         =   "Asignar Lotes Automáticamente"
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
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   1920
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.TextBox txtLoteInterno 
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
         TabIndex        =   6
         Text            =   "0"
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox txtExistenciaLote 
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
         Left            =   8760
         TabIndex        =   5
         Text            =   "0"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtCanBonif 
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
         Left            =   8400
         TabIndex        =   4
         Text            =   "0"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton optDescBonif 
         Caption         =   "Descuento por Und Bonificada"
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   1680
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.OptionButton optDescPorc 
         Caption         =   "Descuento Porcentual sin Bonificación"
         Height          =   255
         Left            =   3360
         TabIndex        =   2
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtCantTotal 
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
         Left            =   8400
         TabIndex        =   1
         Text            =   "0"
         Top             =   1680
         Width           =   1215
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
         Left            =   240
         TabIndex        =   40
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Total :"
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
         TabIndex        =   39
         Top             =   2520
         Width           =   615
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
         Left            =   4800
         TabIndex        =   38
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Precio"
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
         TabIndex        =   37
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "% Impuesto :"
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
         Left            =   2280
         TabIndex        =   36
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Line Line1 
         X1              =   480
         X2              =   9600
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label9 
         Caption         =   "SubTotal"
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
         TabIndex        =   35
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label lblTotalImpuesto 
         Caption         =   "Impuesto :"
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
         Left            =   4680
         TabIndex        =   34
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Lote"
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
         Left            =   3600
         TabIndex        =   33
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "Descuento"
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
         Left            =   2400
         TabIndex        =   32
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Existencia Lote :"
         Height          =   255
         Left            =   7560
         TabIndex        =   31
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "CantBonif :"
         Height          =   255
         Left            =   7320
         TabIndex        =   30
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblCantidadTotal 
         Caption         =   "Cantidad Total :"
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
         Left            =   6840
         TabIndex        =   29
         Top             =   1680
         Width           =   1455
      End
   End
   Begin TrueOleDBGrid60.TDBGrid TDBGFAC 
      Height          =   2895
      Left            =   240
      OleObjectBlob   =   "frmPedidoFactura.frx":2DF8
      TabIndex        =   69
      Top             =   2100
      Width           =   14535
   End
   Begin MSComctlLib.StatusBar StatusSubTotal 
      Height          =   495
      Left            =   10440
      TabIndex        =   70
      Top             =   5220
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
      Left            =   10440
      TabIndex        =   71
      Top             =   7380
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
   Begin MSComctlLib.StatusBar StatusImpuesto 
      Height          =   495
      Left            =   10440
      TabIndex        =   72
      Top             =   6660
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
   Begin MSComctlLib.StatusBar StatusDescuento 
      Height          =   495
      Left            =   10440
      TabIndex        =   73
      Top             =   5940
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
   Begin VB.Image imgEdit 
      Height          =   480
      Left            =   14880
      Picture         =   "frmPedidoFactura.frx":809C
      Top             =   5100
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgAdd 
      Height          =   480
      Left            =   14880
      Picture         =   "frmPedidoFactura.frx":84DE
      Top             =   5940
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgOk 
      Height          =   480
      Left            =   14880
      Picture         =   "frmPedidoFactura.frx":8920
      Top             =   6660
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmPedidoFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsttmpProdFac As ADODB.Recordset ' para la fuente del grid ( Productos )
Dim gModeEdit As Boolean
Dim gModeAdd As Boolean
Dim gConfirmado As Boolean ' se setea a true cuando el usuario acepta un item de la factura para el grid
Dim gSaveChange As Boolean
Dim iCantBodegasFacturables As Integer ' Bodegas facturables del usuario
Dim gbLoteInProcess As Boolean

Private Sub chkAplicaBonificacion_Click()
If chkAplicaBonificacion.value = 1 Then
    cmdEscala.Enabled = True
Else
    cmdEscala.Enabled = False
End If
End Sub

Private Sub chkLoteAutomaticos_Click()
If chkLoteAutomaticos.value = 0 Then
    txtIdLote.Text = ""
    txtLoteInterno.Text = ""
    txtExistenciaLote.Text = "0"
End If
End Sub

Private Sub cmdAddItem_Click()
ClearControlsLinea
gConfirmado = False
gModeAdd = True
gModeEdit = False
gSaveChange = False
cmdProducto.SetFocus
End Sub
Private Sub ClearControlsLinea()
txtCodProd.Text = ""
txtDescProd.Text = ""
txtPrecio.Text = "0"
txtPorcImpuesto.Text = "0"
txtImpuesto.Text = "0"
txtCantidad.Text = ""
txtSubTotal.Text = "0"
txtTotalImpuesto.Text = "0"
txtTotal.Text = "0"
txtIdLote.Text = ""
txtLoteInterno.Text = ""
txtExistenciaLote.Text = 0
End Sub

Private Sub TotalizaDetalle()
Dim bmPos As Variant
Dim dTotalSubTotal As Double
Dim dTotalImpuesto As Double
Dim dTotal As Double

If Not rsttmpProdFac.EOF Then
    bmPos = rsttmpProdFac.Bookmark
    rsttmpProdFac.MoveFirst
    dTotalSubTotal = 0
    dTotalImpuesto = 0
    dTotal = 0
    While Not rsttmpProdFac.EOF
    
         dTotalSubTotal = dTotalSubTotal + rsttmpProdFac("SubTotal").value
         dTotalImpuesto = dTotalImpuesto + rsttmpProdFac("TotalImpuesto").value
         dTotal = dTotal + rsttmpProdFac("Total").value
    rsttmpProdFac.MoveNext
    Wend
        rsttmpProdFac.Bookmark = bmPos
Else
        dTotalSubTotal = 0
        dTotalImpuesto = 0
        dTotal = 0
End If
        TDBGFAC.Columns("SubTotal").NumberFormat = "###,###,##0.#0"
        TDBGFAC.Columns("SubTotal").FooterText = Format(dTotalSubTotal, "###,###,##0.#0")
        TDBGFAC.Columns("TotalImpuesto").NumberFormat = "###,###,##0.#0"
        TDBGFAC.Columns("TotalImpuesto").FooterText = Format(dTotalImpuesto, "###,###,##0.#0")
        TDBGFAC.Columns("Total").NumberFormat = "###,###,##0.#0"
        TDBGFAC.Columns("Total").FooterText = Format(dTotal, "###,###,##0.#0") '* CDbl(txtCantidad.Text)   'txtTotal.Caption
        
        StatusSubTotal.Panels(2).Text = Format(dTotalSubTotal, "###,###,##0.#0")  'TDBGFAC.Columns("TotalLinea").FooterText
        StatusImpuesto.Panels(2).Text = Format(dTotalImpuesto, "###,###,##0.#0")
        StatusTotal.Panels(2).Text = Format(dTotal, "###,###,##0.#0")  '

End Sub

Private Sub cmdBodega_Click()
   Dim frm As frmBrowseCat

    Set frm = New frmBrowseCat
    frm.gsCaptionfrm = "Bodegas"
    frm.gsTablabrw = "vinvBodegaUsuario"
    frm.gsCodigobrw = "IDBodega"
    frm.gbTypeCodeStr = False
    frm.gsDescrbrw = "DESCRBODEGA"
    frm.gbFiltra = True
    frm.gsFiltro = "USUARIO='" & gsUSUARIO & "' AND FACTURA =1"
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
        txtcodBodega.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      txtDescrBodega.Text = frm.gsDescrbrw
      fmtTextbox txtDescrBodega, "R"
      cmdVendedor.SetFocus
    End If
End Sub

Private Sub cmdCliente_Click()
   Dim frm As frmBrowseCat

    Set frm = New frmBrowseCat
    frm.gsCaptionfrm = "Clientes"
    frm.gsTablabrw = "ccCLIENTE"
    frm.gsCodigobrw = "IDCLIENTE"
    frm.gbTypeCodeStr = False
    frm.gsDescrbrw = "Nombre"
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
        txtCodCliente.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      txtNombre.Text = frm.gsDescrbrw
      fmtTextbox txtNombre, "R"
      cmdTipo.SetFocus
    End If
End Sub



Private Sub cmdDelItem_Click()
Dim lbok As Boolean
On Error GoTo error
If Not rsttmpProdFac.EOF Then
  If Mensaje("Está seguro que desea eliminar el item seleccionado " & Chr(13) & rsttmpProdFac!Descr, ICO_PREGUNTA, True) = True Then
    rsttmpProdFac.Delete
    rsttmpProdFac.MoveFirst
    If rsttmpProdFac.EOF Then
        GSSQL = gsCompania & ".fafGetDetallePedidoToGrid -1, -1,-1"
        If rsttmpProdFac.State = adStateOpen Then rsttmpProdFac.Close
        Set rsttmpProdFac = GetRecordset(GSSQL)
        ClearControlsLinea
        cmdDelItem.Enabled = False
        cmdEditItem.Enabled = False
        cmdSave.Enabled = False
        gModeEdit = False
        gModeAdd = False
        gConfirmado = False
        gSaveChange = False
    End If
        TotalizaDetalle
    
  End If
End If
Exit Sub
error:
    lbok = Mensaje("Hubo un error al eliminar un item del Pedido : " & err.Description, ICO_ERROR, False)
End Sub

Private Sub cmdEditItem_Click()
Dim lbok As Boolean
Dim sTotal As String
Dim sTotalD As String
Dim dTotal As Double

If Not Val_TextboxNum(txtCantidad) Then
    lbok = Mensaje("La Cantidad debe ser numérica", ICO_ERROR, False)
    Exit Sub
End If
'-- esto va a ser para facturacion
'    ExistenciaDisp = ExistenciaDisponible(gparametros.BodFacturacion, txtCodProd.Text)      ' si es producto
'    If ExistenciaDisp < txtCantidad.Text Then
'      lbOk = Mensaje("Esa cantidad no es cubierta con el inventario existente. Tiene que revisar el inventario. " & _
'        " El sistema le va a dejar facturar para no perder el cliente, hay problemas con el inventario y debe chequearlo. de ese producto solo existen : " & ExistenciaDisp, ICO_ADVERTENCIA, False)
'        Exit Sub
'    End If
  If Not rsttmpProdFac.EOF Then
    rsttmpProdFac("CantidadPedida").value = txtCantidad.Text
    rsttmpProdFac("SubTotal").value = txtSubTotal.Text
    rsttmpProdFac("TotalImpuesto").value = txtTotalImpuesto.Text
    rsttmpProdFac("Total").value = txtTotal.Text
    rsttmpProdFac.Update
    TotalizaDetalle
  End If
End Sub

Private Sub cmdEscala_Click()
Dim sPorCada As String
Dim sBonifica As String
Dim lbok As Boolean
If Not Val_TextboxNum(txtCantidad) Then
    lbok = Mensaje("Para seleccionar una escala de Bonificación, la Cantidad no puede quedar vacía", ICO_ADVERTENCIA, False)
    Exit Sub
End If
Dim frm As New frmEscalaBonificacion
Set frm = New frmEscalaBonificacion
frm.gsFormCaption = "Escalas de Bonificación"
frm.gsTitle = "Escalas de Bonificación"
frm.gsIDProducto = txtCodProd.Text
frm.gsDescr = txtDescProd.Text
frm.giCantidadFuente = CInt(txtCantidad.Text)
frm.gbOnlyShow = True
frm.Show vbModal
sPorCada = frm.gsPorCada
sBonifica = frm.gsBonifica
If sPorCada = "" Or sPorCada = "0" Or sBonifica = "" Or sBonifica = "0" Then
lbok = Mensaje("Ud no seleccionó una Escala de Bonificación", ICO_ERROR, False)
Exit Sub
End If
txtFAPorCada.Text = sPorCada
txtFABonifica.Text = sBonifica
RefreshCantidad
Set frm = Nothing
End Sub

Private Sub cmdLote_Click()
   Dim frm As frmBrowseCat

    Set frm = New frmBrowseCat
    frm.gsCaptionfrm = "Lotes"
    frm.gsTablabrw = "vinvExistenciaLote"
    frm.gsCodigobrw = "IDLote"
    frm.gbTypeCodeStr = False
    frm.gsDescrbrw = "LoteProveedor"
    frm.gsMuestraExtra = "SI"
    frm.gsFieldExtrabrw = "FechaVencimiento"
    frm.gsMuestraExtra2 = "SI"
    frm.gsFieldExtrabrw2 = "Existencia"
    frm.gbFiltra = True
    frm.gsFiltro = "IDBodega=" & txtcodBodega.Text & " and IDProducto=" & txtCodProd.Text & " and Existencia>0"
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
        txtIdLote.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      txtLoteInterno.Text = frm.gsDescrbrw
      fmtTextbox txtLoteInterno, "R"
      txtExistenciaLote.Text = frm.gsExtraValor2
    End If
End Sub

Private Sub cmdOkCO_Click()
Dim lbok As Boolean
Dim sFiltro As String
Dim ExistenciaDisp As Double
Dim sFecha As String
Dim rst As ADODB.Recordset
Dim frmAuto As frmAutoSugiereLotes
Dim dTotalSugeridoLote As Double
gbLoteInProcess = False

    If Not ValCtrls Then
        'lbOk = Mensaje("Revise sus datos por favor !!! " & gsOperacionError, ICO_ADVERTENCIA, False)
        Exit Sub
    End If
    sFecha = Format(Str(DTPFecFac.value), "yyyymmdd")
    
    If chkLoteAutomaticos.value = 1 Then
        gbLoteInProcess = True
        Set rst = New ADODB.Recordset
        If rst.State = adStateOpen Then grst.Close
        rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
        rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
        rst.CursorLocation = adUseClient ' Cursor local al cliente
        rst.LockType = adLockOptimistic
        Set frmAuto = New frmAutoSugiereLotes
        frmAuto.gsIDBodega = txtcodBodega.Text
        frmAuto.gsIDProducto = txtCodProd.Text
        frmAuto.gdCantidad = CDbl(txtCantidad.Text)
        frmAuto.gsDescrProducto = txtDescProd.Text
        frmAuto.gsDescrBodega = Me.txtDescrBodega.Text
        frmAuto.gsFormCaption = "Asignación de Lotes"
        frmAuto.gsTitle = "Asinación de Lotes"
        dTotalSugeridoLote = frmAuto.getTotalSugeridoporLote()
        If dTotalSugeridoLote < frmAuto.gdCantidad Then
            lbok = Mensaje("Las existencias disponibles en lote son " & Str(dTotalSugeridoLote), ICO_ADVERTENCIA, False)
            Set rst = Nothing
            Set frmAuto = Nothing
        Else
            frmAuto.Show vbModal
            Set rst = frmAuto.grst
        End If
       
        ' llenar el recordset de facturacion con los datos del lote
        If rst Is Nothing Then Exit Sub
        rst.MoveFirst
        While Not rst.EOF
        
            If Not rsttmpProdFac.EOF Then
              rsttmpProdFac.MoveLast
            End If
            rsttmpProdFac.AddNew
            rsttmpProdFac("IDBodega").value = txtcodBodega.Text
            rsttmpProdFac("IDVendedor").value = txtCodVendedor.Text
            rsttmpProdFac("IDCliente").value = txtCodCliente.Text
            rsttmpProdFac("Fecha").value = Format(DTPFecFac.value, "yyyy-mm-dd") 'DTPFecFac.value
            rsttmpProdFac("IDProducto").value = txtCodProd.Text
            rsttmpProdFac("IDLote").value = rst("IDLote").value
            rsttmpProdFac("Descr").value = txtDescProd.Text
            rsttmpProdFac("CantidadPedida").value = rst("Cantidad").value 'txtCantidad.Text
            rsttmpProdFac("PorcImpuesto").value = txtPorcImpuesto.Text
            rsttmpProdFac("Impuesto").value = CDbl(txtPrecio.Text) * (CDbl(txtPorcImpuesto.Text) / 100) 'txtImpuesto.Text
            rsttmpProdFac("SubTotal").value = CDbl(txtPrecio.Text) * rst("Cantidad").value
            rsttmpProdFac("TotalImpuesto").value = CDbl(txtPrecio.Text) * rst("Cantidad").value * (CDbl(txtPorcImpuesto.Text) / 100)
            rsttmpProdFac("Total").value = CDbl(txtPrecio.Text) * rst("Cantidad").value + CDbl(txtPrecio.Text) * rst("Cantidad").value * (CDbl(txtPorcImpuesto.Text) / 100)
            rsttmpProdFac("PrecioFarmaciaLocal").value = txtPrecio.Text
        rst.MoveNext
        Wend
        gbLoteInProcess = False
        '-*********************************************************
        
    Else ' El usuario está asignando el lote manualmente
        If ExiteRstKey(rsttmpProdFac, "IDProducto=" & txtCodProd.Text & " AND IDBODEGA=" & txtcodBodega.Text & " and IDLOTE=" & txtIdLote.Text) Then
          lbok = Mensaje("Ya Existe ese Producto. ", ICO_ERROR, False)
          Exit Sub
        
        End If
        If CDbl(txtCantidad.Text) > CDbl(txtExistenciaLote.Text) Then
          lbok = Mensaje("La Cantidad requerida no es satisfecha por la existencia en ese lote ", ICO_ERROR, False)
          Exit Sub
            
        End If
        
        If Not rsttmpProdFac.EOF Then
          rsttmpProdFac.MoveLast
        End If
        rsttmpProdFac.AddNew
        rsttmpProdFac("IDBodega").value = txtcodBodega.Text
        rsttmpProdFac("IDVendedor").value = txtCodVendedor.Text
        rsttmpProdFac("IDCliente").value = txtCodCliente.Text
        rsttmpProdFac("Fecha").value = Format(DTPFecFac.value, "yyyy-mm-dd") 'DTPFecFac.value
        rsttmpProdFac("IDProducto").value = txtCodProd.Text
        rsttmpProdFac("IDLote").value = txtIdLote.Text
        
        rsttmpProdFac("Descr").value = txtDescProd.Text
        rsttmpProdFac("CantidadPedida").value = txtCantidad.Text
        rsttmpProdFac("Impuesto").value = txtImpuesto.Text
        rsttmpProdFac("PorcImpuesto").value = txtPorcImpuesto.Text
        rsttmpProdFac("SubTotal").value = txtSubTotal.Text
        rsttmpProdFac("TotalImpuesto").value = txtTotalImpuesto.Text
        rsttmpProdFac("Total").value = txtTotal.Text
        rsttmpProdFac("PrecioFarmaciaLocal").value = txtPrecio.Text
    
             
    End If

'    If ExiteItem(rsttmpProdFac, "IDProducto", txtCodProd.Text) Then
'      lbOk = Mensaje("Ya Existe ese Producto. ", ICO_ERROR, False)
'      Exit Sub
'    End If
    

  gConfirmado = True
  ' Actualizando la tabla detalle de transaccion compra
  cmdSave.Enabled = True
  'cmdCobro.Visible = True
  'cmdCobro.Enabled = True
    
    
  TotalizaDetalle
   Set TDBGFAC.DataSource = rsttmpProdFac
   TDBGFAC.Refresh
  cmdAddItem.Enabled = True
  cmdDelItem.Enabled = True
  imgAdd.Visible = True
  cmdAddItem.SetFocus
  imgOk.Visible = False
  
  cmdOkCO.Enabled = False
  gModeAdd = True ' ANTES ERA FALSE
  gModeEdit = False
  gSaveChange = True
  
  cmdAddItem.Enabled = True
  cmdDelItem.Enabled = True
  imgAdd.Visible = True
  cmdAddItem.SetFocus
  imgOk.Visible = False

  gDescuentoProm = 0
  gDescuentoExtra = 0



  
End Sub

Private Sub cmdProducto_Click()
Dim sResultado As String
Dim frm As frmBrowseCat
Dim sImpuesto As String
Dim sPrecio As String
Dim dictValue As New Dictionary

    Set frm = New frmBrowseCat
    frm.gsCaptionfrm = "Productos"
    frm.gsTablabrw = "vinvProducto"
    frm.gsCodigobrw = "IDProducto"
    frm.gbTypeCodeStr = False
    frm.gsDescrbrw = "Descr"
    'frm.gsExtraValor1 = "SI"
    frm.gsMuestraExtra = "SI"
    frm.gsFieldExtrabrw = "PORCIMPUESTO"
    frm.gsMuestraExtra2 = "NO"
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
        txtCodProd.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      txtDescProd.Text = frm.gsDescrbrw
        'lbOk = getValueFieldFromTable("vinvProducto", "PorcImpuesto", "IDProducto=" & txtCodProd.Text, "PrecioFarmaciaLocal", sImpuesto, sPrecio, True, True)
        lbok = getValueFieldsFromTable("vinvProducto", "PrecioFarmaciaLocal,PorcImpuesto", "IDProducto=" & txtCodProd.Text, dictValue)
        If Not lbok Then
            lbok = Mensaje("Hay un error con el porcentaje de Impuesto o el Precio para ese producto", ICO_ERROR, False)
        Else
            txtPrecio.Text = dictValue("PrecioFarmaciaLocal")
            txtPrecio.Text = Format(txtPrecio.Text, "#,##0.#0")
            txtPorcImpuesto.Text = dictValue("PorcImpuesto") 'Trim(sImpuesto)
            txtPorcImpuesto.Text = Format(txtPorcImpuesto.Text, "#,##0.#0")
            txtImpuesto.Text = CDbl(txtPrecio.Text) * (CDbl(txtPorcImpuesto.Text) / 100)
            txtImpuesto.Text = Format(txtImpuesto.Text, "#,##0.#0")
            txtFAPorCada.Text = "0"
            txtFABonifica.Text = "0"
            
        End If

      
      fmtTextbox txtDescProd, "R"
      txtCantidad.SetFocus
    End If
    Set frm = Nothing
End Sub

Private Sub cmdSave_Click()
Dim lbok As Boolean
Dim sFecha As String
Dim sIDPedido As String

    If Not ValCtrls Then
        lbok = Mensaje("Revise sus datos por favor !!! ", ICO_ADVERTENCIA, False)
        Exit Sub
    End If
     If rsttmpProdFac.RecordCount > 0 Then
     sFecha = Format(Str(DTPFecFac.value), "yyyymmdd")
        
        gConet.BeginTrans
        lbok = fafUpdatePedido("I", sIDPedido, txtcodBodega.Text, txtCodCliente.Text, txtCodVendedor.Text, sFecha, "0", "0", "0")
        lblNoPedido.Caption = sIDPedido
        rsttmpProdFac.MoveFirst
        While Not rsttmpProdFac.EOF And lbok
            lbok = fafUpdatePedidoLinea("I", sIDPedido, rsttmpProdFac("IDBodega").value, rsttmpProdFac("IDCliente").value, rsttmpProdFac("IDVendedor").value, sFecha, _
            rsttmpProdFac("IDProducto").value, rsttmpProdFac("CantidadPedida").value, rsttmpProdFac("PrecioFarmaciaLocal").value, rsttmpProdFac("SubTotal").value, _
            rsttmpProdFac("TotalImpuesto").value, rsttmpProdFac("Total").value)
            rsttmpProdFac.MoveNext
        Wend
        If lbok = False Then
            gConet.RollbackTrans
        Else
            gConet.CommitTrans
            GSSQL = gsCompania & ".fafGetDetallePedidoToGrid " & txtcodBodega.Text & "," & txtCodCliente.Text & "," & sIDPedido
                 
            If rsttmpProdFac.State = adStateOpen Then rsttmpProdFac.Close
            Set rsttmpProdFac = GetRecordset(GSSQL)
            Set TDBGFAC.DataSource = rsttmpProdFac
            TDBGFAC.Refresh
            ImprimePedido rsttmpProdFac, False
            Unload Me
'            cmdDelItem.Enabled = False
'            cmdEditItem.Enabled = False
'            cmdSave.Enabled = False
'            gModeEdit = False
'            gModeAdd = False
'            gConfirmado = False
'            gSaveChange = False
            ' Inhabilita los iconos
        End If
     End If
End Sub

Private Sub cmdTipo_Click()
   Dim frm As frmBrowseCat

    Set frm = New frmBrowseCat
    frm.gsCaptionfrm = "Tipo Factura"
    frm.gsTablabrw = "vfacTipoFactura"
    frm.gsCodigobrw = "Codigo"
    frm.gbTypeCodeStr = False
    frm.gsDescrbrw = "Descr"
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
        txtTipo.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      txtDescrTipo.Text = frm.gsDescrbrw
      fmtTextbox txtDescrTipo, "R"
      cmdProducto.SetFocus
    End If
End Sub

Private Sub cmdVendedor_Click()
   Dim frm As frmBrowseCat

    Set frm = New frmBrowseCat
    frm.gsCaptionfrm = "Vendedor"
    frm.gsTablabrw = "fafVendedor"
    frm.gsCodigobrw = "IDVendedor"
    frm.gbTypeCodeStr = False
    frm.gsDescrbrw = "Nombre"
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
        txtCodVendedor.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      txtDescrVendedor.Text = frm.gsDescrbrw
      fmtTextbox txtDescrVendedor, "R"
      cmdCliente.SetFocus
    End If
End Sub

Private Sub Form_Load()
Dim lbok As Boolean
Dim sResultado1 As String
Dim sResultado2 As String

iCantBodegasFacturables = fafgetCantBodegaFacturableForUser(gsUSUARIO)
If iCantBodegasFacturables = 0 Then
    lbok = Mensaje("Ud no tiene asignada ninguna bodega facturable, por favor vea al administrador del Sistema", ICO_ERROR, False)
    'Unload Me
    Exit Sub
Else
    If iCantBodegasFacturables = 1 Then
        lbok = getValueFieldFromTable("vinvBodegaUsuario", "IDBODEGA", "USUARIO='" & gsUSUARIO & "' AND FACTURA=1", "DESCRBODEGA", sResultado1, sResultado2, True, False)
        txtcodBodega.Text = sResultado1
        txtDescrBodega.Text = sResultado2
        fmtTextbox txtcodBodega, "R"
        fmtTextbox txtDescrBodega, "R"
        cmdBodega.Enabled = False
        'cmdVendedor.SetFocus
    End If
End If
gModeEdit = False
gModeAdd = False
gSaveChange = False
gbLoteInProcess = False
DTPFecFac.value = Now
PreparaRst
fmtTextbox txtTotal, "R"
fmtTextbox txtImpuesto, "R"
fmtTextbox txtTotalImpuesto, "R"
fmtTextbox txtSubTotal, "R"
fmtTextbox txtPrecio, "R"
fmtTextbox txtPorcImpuesto, "R"
fmtTextbox txtIdLote, "R"
fmtTextbox txtExistenciaLote, "R"
fmtTextbox txtLoteInterno, "R"
SetColumnSizeGrid
Set TDBGFAC.DataSource = rsttmpProdFac
TDBGFAC.Refresh

End Sub

Private Sub SetColumnSizeGrid()
TDBGFAC.Columns("IDProducto").Width = 1110.047
TDBGFAC.Columns("Descr").Width = 4000
TDBGFAC.Columns("IDLote").Width = 1200
TDBGFAC.Columns("PRecio").Width = 1454.74
TDBGFAC.Columns("Cantidad").Width = 1574.929
TDBGFAC.Columns("SubTotal").Width = 1785.26
TDBGFAC.Columns("IMpuesto").Width = 1574.929
TDBGFAC.Columns("Total").Width = 1635.024

End Sub

Private Sub PreparaRst()
      ' preparacion del recordset fuente del grid de compra
      ' recordar que este recordset va a ser temporal, no se hara addnew a la bd
      ' lleva además de los campos de la tabla detalle de compra, la descripcion del producto
      Set rsttmpProdFac = New ADODB.Recordset
      If rsttmpProdFac.State = adStateOpen Then rsttmpProdFac.Close
      rsttmpProdFac.ActiveConnection = gConet 'Asocia la conexión de trabajo
      rsttmpProdFac.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
      rsttmpProdFac.CursorLocation = adUseClient ' Cursor local al cliente
      rsttmpProdFac.LockType = adLockOptimistic
        
      GSSQL = gsCompania & ".fafGetDetallePedidoToGrid -1, -1,-1"
           
      If rsttmpProdFac.State = adStateOpen Then rsttmpProdFac.Close
      Set rsttmpProdFac = GetRecordset(GSSQL)
End Sub

Private Function ValCtrls() As Boolean
Dim lbok As Boolean
Dim sDescr As String
On Error GoTo salir
gsOperacionError = ""
lbok = True

If Format(DTPFecFac.value, "yyyy-mm-dd") < Format(Now, "yyyy-mm-dd") Then
lbok = Mensaje("La fecha del Pedido es menor que la fecha actual, desea corregirla ", ICO_PREGUNTA, True)
    If lbok = True Then
        lbok = False
        DTPFecFac.SetFocus
        GoTo salir
    Else
        lbok = True
    End If
End If

If Format(DTPFecFac.value, "yyyy-mm-dd") > Format(Now, "yyyy-mm-dd") Then
lbok = Mensaje("La fecha del Pedido es mayor que la fecha actual, desea corregirla ", ICO_PREGUNTA, True)
    If lbok = True Then
        lbok = False
        DTPFecFac.SetFocus
        GoTo salir
    Else
        lbok = True
    End If
End If

If Not Val_TextboxNum(txtcodBodega) Then
 gsOperacionError = "El código de la Bodega debe ser numérico."
 lbok = Mensaje(gsOperacionError, ICO_ADVERTENCIA, False)
 lbok = False
 cmdBodega.SetFocus
 'txtcodBodega.SetFocus
 GoTo salir
End If

If txtcodBodega.Text <> "" Then
    sDescr = GetDescrCat("IDBodega", txtcodBodega.Text, "invBodega", "Descr")
  If sDescr = "" Then
    gsOperacionError = "La Bodega No Existe."
    lbok = Mensaje(gsOperacionError, ICO_ADVERTENCIA, False)
    txtcodBodega.Text = ""
    txtDescrBodega.Text = ""
    lbok = False
    'txtcodBodega.SetFocus
    cmdBodega.SetFocus
    
    GoTo salir
  Else
    txtDescrBodega.Text = sDescr
  End If
End If


If Not Val_TextboxNum(txtCodVendedor) Then
 gsOperacionError = "El código del Vendedor debe ser numérico."
 lbok = Mensaje(gsOperacionError, ICO_ADVERTENCIA, False)
 lbok = False
 'txtCodVendedor.SetFocus
 cmdVendedor.SetFocus
 GoTo salir
End If


If txtCodVendedor.Text <> "" Then
    sDescr = GetDescrCat("IDVENDEDOR", txtCodVendedor.Text, "FAFVENDEDOR", "Nombre")
  If sDescr = "" Then
    gsOperacionError = "El Vendedor no existe."
    lbok = Mensaje(gsOperacionError, ICO_ADVERTENCIA, False)
    txtCodVendedor.Text = ""
    txtDescrVendedor.Text = ""
    lbok = False
    'txtCodVendedor.SetFocus
    cmdVendedor.SetFocus
 
    GoTo salir
  Else
    txtDescrVendedor.Text = sDescr
  End If
End If

If Not Val_TextboxNum(txtCodCliente) Then
 gsOperacionError = "El código del Cliente debe ser numérico."
 lbok = Mensaje(gsOperacionError, ICO_ADVERTENCIA, False)
 lbok = False
 'txtCodCliente.SetFocus
 cmdCliente.SetFocus
 GoTo salir
End If

If txtCodCliente.Text <> "" Then
    sDescr = GetDescrCat("IDCLIENTE", txtCodCliente.Text, "ccCLIENTE", "Nombre")
  If sDescr = "" Then
    gsOperacionError = "El CLIENTE no existe."
    lbok = Mensaje(gsOperacionError, ICO_ADVERTENCIA, False)
    txtCodCliente.Text = ""
    txtNombre.Text = ""
    lbok = False
    'txtCodCliente.SetFocus
    cmdCliente.SetFocus
    
    GoTo salir
  Else
    txtDescrVendedor.Text = sDescr
  End If
End If

If txtTipo.Text = "" Then
    gsOperacionError = "El Tipo de la Factura debe indicarse..."
    lbok = Mensaje(gsOperacionError, ICO_ADVERTENCIA, False)
    lbok = False
    cmdTipo.SetFocus
    GoTo salir
End If

If txtTipo.Text <> "" Then
    sDescr = GetDescrCat("Codigo", txtTipo.Text, "[vfacTipoFactura]", "Descr")

  If sDescr = "" Then
    gsOperacionError = "El Tipo de Factura no existe."
    lbok = Mensaje(gsOperacionError, ICO_ADVERTENCIA, False)
    txtTipo.Text = ""
    txtDescrTipo.Text = ""
    'txtTipo.SetFocus
    lbok = False
    cmdTipo.SetFocus
    
    GoTo salir
  Else
    txtDescrTipo.Text = sDescr
  End If
End If



If Not Val_TextboxNum(txtCodProd) Then
 lbok = False
 gsOperacionError = "El código del Producto debe ser numérico."
 lbok = Mensaje(gsOperacionError, ICO_ADVERTENCIA, False)
    lbok = False
    txtCodProd.Text = ""
    txtDescProd.Text = ""
    cmdProducto.SetFocus
 'txtCodProd.SetFocus
 GoTo salir
End If


If txtCodProd.Text <> "" Then
    sDescr = GetDescrCat("IDProducto", txtCodProd.Text, "invPRODUCTO", "Descr")

  If sDescr = "" Then
    gsOperacionError = "El Producto no existe."
    lbok = Mensaje(gsOperacionError, ICO_ADVERTENCIA, False)
    txtCodProd.Text = ""
    txtDescProd.Text = ""
    'txtCodProd.SetFocus
    lbok = False
    cmdProducto.SetFocus
    
    GoTo salir
  Else
    txtDescProd.Text = sDescr
  End If
End If


  If txtCantidad.Text = "" Then
    gsOperacionError = "Ud no ha digitado la Cantidad"
    lbok = Mensaje(gsOperacionError, ICO_ADVERTENCIA, False)
    lbok = False
    txtCantidad.SetFocus
    GoTo salir
  End If
  
If Not Val_TextboxNum(txtCantidad) Then
 lbok = False
  gsOperacionError = "La Cantidad debe ser numérica."
  lbok = Mensaje(gsOperacionError, ICO_ADVERTENCIA, False)
  lbok = False
 txtCantidad.SetFocus
 GoTo salir
End If
If chkLoteAutomaticos.value = 0 Then
      If txtIdLote.Text = "" Then
        gsOperacionError = "Ud no ha seleccionado el lote"
        lbok = Mensaje(gsOperacionError, ICO_ADVERTENCIA, False)
        lbok = False
        cmdLote.SetFocus
        GoTo salir
      End If
    
    
    If Not Val_TextboxNum(txtIdLote) Then
     lbok = False
     gsOperacionError = "El código del Lote debe ser numérico."
     lbok = Mensaje(gsOperacionError, ICO_ADVERTENCIA, False)
        lbok = False
        txtIdLote.Text = ""
        txtLoteInterno.Text = ""
        txtExistenciaLote.Text = "0"
        cmdLote.SetFocus
     'txtCodProd.SetFocus
     GoTo salir
    End If
End If



If Not Val_TextboxNum(txtPrecio) Then
 gsOperacionError = "El Precio del Producto debe ser numérico."
 lbok = Mensaje(gsOperacionError, ICO_ADVERTENCIA, False)
 lbok = False
 txtPrecio.SetFocus
 GoTo salir
End If
lbok = True
ValCtrls = lbok
Exit Function
salir:

ValCtrls = lbok
End Function


Private Sub qw21_Click()

End Sub

Private Sub TDBGFAC_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'If gSaveChange = False And gModeAdd = True Then
'  Exit Sub
'End If

'If gConfirmado = True And gModeAdd = True And gModeEdit = False Then
'  gModeAdd = True
'  cmdAddItem_Click
'  Exit Sub
'End If

If (rsttmpProdFac Is Nothing) Or gbLoteInProcess Then Exit Sub
If Not rsttmpProdFac.EOF Then
    GetDataFromGridToControl
    gModeEdit = True
    gModeAdd = False
    cmdEditItem.Enabled = True
    cmdAddItem.Enabled = True
    
End If
End Sub
Private Sub GetDataFromGridToControl()
If Not rsttmpProdFac.EOF Then
    txtCodProd.Text = rsttmpProdFac("IDProducto").value
    txtDescProd.Text = rsttmpProdFac("Descr").value
    txtCantidad.Text = rsttmpProdFac("CantidadPedida").value
    txtPrecio.Text = rsttmpProdFac("PrecioFarmaciaLocal").value
  txtImpuesto.Text = rsttmpProdFac("Impuesto").value
  txtPorcImpuesto.Text = rsttmpProdFac("PorcImpuesto").value
  txtSubTotal.Text = rsttmpProdFac("SubTotal").value
  txtTotalImpuesto.Text = rsttmpProdFac("TotalImpuesto").value
  txtTotal.Text = rsttmpProdFac("Total").value
  txtIdLote.Text = rsttmpProdFac("IDLote").value
End If
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    RefreshCantidad
End If
End Sub

Private Sub RefreshCantidad()
    If txtCantidad.Text <> "" Then

            If Not Val_TextboxNum(txtCantidad) Then
              lbok = Mensaje("Digite un valor correcto por favor ", ICO_ADVERTENCIA, False)
              txtCantidad.SetFocus
              Exit Sub
            End If
            txtCanBonif.Text = getUnidadesBonificadas(CDbl(txtCantidad.Text), CDbl(txtFAPorCada.Text), CDbl(txtFABonifica.Text))
            txtCantTotal.Text = Val(txtCantidad.Text) + Val(txtCanBonif.Text)
            TotalDescuento.Text = Format(CDbl(txtCanBonif.Text) * CDbl(txtPrecio.Text), "#,##0.#0")
            If gModeAdd = False And gModeEdit = True Then
              
              cmdEditItem.Enabled = True
              imgAdd.Visible = False
              imgEdit.Visible = True
              ShowTotalesLinea
              cmdEditItem.SetFocus
              Exit Sub
            End If
            ShowTotalesLinea

            cmdOkCO.Enabled = True
            imgOk.Visible = True
            cmdOkCO.SetFocus
    End If

End Sub


Private Sub ShowTotalesLinea()
            txtTotal.Text = CalculaTotalLineaVenta(CDbl(txtCantidad.Text), CDbl(txtPrecio.Text), CDbl(txtPorcImpuesto.Text))
            txtTotal.Text = Format(txtTotal.Text, "#,##0.#0")
            txtImpuesto.Text = CDbl(txtPrecio.Text) * (CDbl(txtPorcImpuesto.Text) / 100)
            txtImpuesto.Text = Format(txtImpuesto.Text, "#,##0.#0")
            txtTotalImpuesto.Text = ((CDbl(txtCantidad.Text) * CDbl(txtPrecio.Text) - CDbl(txtCanBonif.Text) * CDbl(txtPrecio.Text)) * CDbl(txtPorcImpuesto.Text) / 100)
            txtTotalImpuesto.Text = Format(txtTotalImpuesto.Text, "#,##0.#0")
            txtSubTotal.Text = (CDbl(txtCantidad.Text) * CDbl(txtPrecio.Text))
            txtSubTotal.Text = Format(txtSubTotal.Text, "#,##0.#0")

End Sub
Private Function CalculaTotalLineaVenta(Cantidad As Double, Precio As Double, PorcImpuesto As Double) As Double
CalculaTotalLineaVenta = (Cantidad * Precio) + (Cantidad * Precio) * PorcImpuesto / 100
End Function

Private Sub txtCantidad_LostFocus()
RefreshCantidad
End Sub

Private Sub txtcodBodega_KeyPress(KeyAscii As Integer)
Dim sDescr As String
Dim lbok As Boolean
If KeyAscii = vbKeyReturn Then
    sDescr = getDescrCatalogo(txtcodBodega, "IDBodega", "invBodega", "Descr")
    If sDescr <> "" Then
        txtDescrBodega.Text = sDescr
    Else
        lbok = Mensaje("Esa Bodega No Existe", ICO_ERROR, False)
    End If
End If
End Sub

Private Sub txtcodBodega_LostFocus()
Dim sDescr As String
Dim lbok As Boolean

    sDescr = getDescrCatalogo(txtcodBodega, "IDBodega", "invBodega", "Descr")
    If sDescr <> "" Then
        txtDescrBodega.Text = sDescr
    Else
        lbok = Mensaje("Esa Bodega No Existe", ICO_ERROR, False)
    End If

End Sub

Private Sub txtCodCliente_KeyPress(KeyAscii As Integer)
Dim sDescr As String
Dim lbok As Boolean
If KeyAscii = vbKeyReturn Then
    sDescr = getDescrCatalogo(txtCodCliente, "CODCliente", "ccCliente", "Nombre")
    If sDescr <> "" Then
        txtNombre.Text = sDescr
    Else
        lbok = Mensaje("Ese Cliente No Existe", ICO_ERROR, False)
    End If
End If
End Sub

Private Sub txtCodCliente_LostFocus()
Dim sDescr As String
Dim lbok As Boolean

    sDescr = getDescrCatalogo(txtCodCliente, "CodCliente", "ccCliente", "Nombre")
    If sDescr <> "" Then
        txtNombre.Text = sDescr
    Else
        lbok = Mensaje("Ese Cliente No Existe", ICO_ERROR, False)
    End If

End Sub

Private Sub txtCodProd_KeyPress(KeyAscii As Integer)
Dim sDescr As String
Dim lbok As Boolean
Dim sImpuesto As String
Dim sPrecio As String
If KeyAscii = vbKeyReturn Then
    sDescr = getDescrCatalogo(txtCodProd, "IDProducto", "invProducto", "Descr")
    If sDescr <> "" Then
        txtDescProd.Text = sDescr
        lbok = getValueFieldFromTable("vinvProducto", "PorcImpuesto", "IDProducto=" & txtCodProd.Text, "PrecioFarmaciaLocal", sImpuesto, sPrecio, True, True)
        If Not lbok Then
            lbok = Mensaje("Hay un error con el porcentaje de Impuesto o el Precio para ese producto", ICO_ERROR, False)
        Else
            txtPrecio.Text = sPrecio
            txtPorcImpuesto.Text = Trim(sImpuesto)
            txtImpuesto.Text = CDbl(txtPrecio.Text) * (CDbl(txtPorcImpuesto.Text) / 100)
            txtPrecio.Text = Format(txtPrecio.Text, "#,##0.#0")
            txtPorcImpuesto.Text = Format(txtPorcImpuesto.Text, "#,##0.#0")
            txtImpuesto.Text = Format(txtImpuesto.Text, "#,##0.#0")
            txtCantidad.SetFocus
        End If
    Else
        lbok = Mensaje("Ese Producto No Existe", ICO_ERROR, False)
    End If
End If
End Sub

Private Sub txtCodProd_LostFocus()
Dim sDescr As String
Dim sResultado As String
Dim sImpuesto As String
Dim sPrecio As String
Dim lbok As Boolean

    sDescr = getDescrCatalogo(txtCodProd, "IDProducto", "invProducto", "Descr")
    If sDescr <> "" Then
        txtDescProd.Text = sDescr
        lbok = getValueFieldFromTable("vinvProducto", "PorcImpuesto", "IDProducto=" & txtCodProd.Text, "PrecioFarmaciaLocal", sImpuesto, sPrecio, True, True)
        If Not lbok Then
            lbok = Mensaje("Hay un error con el porcentaje de Impuesto o el Precio para ese producto", ICO_ERROR, False)
        Else
            txtPrecio.Text = sPrecio
            txtPorcImpuesto.Text = Trim(sImpuesto)
            txtImpuesto.Text = CDbl(txtPrecio.Text) * (CDbl(txtPorcImpuesto.Text) / 100)
            txtPrecio.Text = Format(txtPrecio.Text, "#,##0.#0")
            txtPorcImpuesto.Text = Format(txtPorcImpuesto.Text, "#,##0.#0")
            txtImpuesto.Text = Format(txtImpuesto.Text, "#,##0.#0")
            txtCantidad.SetFocus
        End If
    End If

End Sub

Private Sub txtCodVendedor_KeyPress(KeyAscii As Integer)
Dim sDescr As String
Dim lbok As Boolean
If KeyAscii = vbKeyReturn Then
    sDescr = getDescrCatalogo(txtCodVendedor, "IDVendedor", "fafVendedor", "Nombre")
    If sDescr <> "" Then
        txtDescrVendedor.Text = sDescr
    Else
        lbok = Mensaje("Ese Vendedor No Existe", ICO_ERROR, False)
    End If
End If
End Sub

Private Sub txtCodVendedor_LostFocus()
Dim sDescr As String
Dim lbok As Boolean

    sDescr = getDescrCatalogo(txtCodVendedor, "IDVendedor", "fafVendedor", "Nombre")
    If sDescr <> "" Then
        txtDescrVendedor.Text = sDescr
    Else
        lbok = Mensaje("Ese Vendedor No Existe", ICO_ERROR, False)
    End If

End Sub

Private Sub txtTipo_KeyPress(KeyAscii As Integer)
Dim sDescr As String
Dim lbok As Boolean
Dim sValor As String
If KeyAscii = vbKeyReturn Then
    sDescr = getDescrCatalogo(txtTipo, "Codigo", "vfacTipoFactura", "Descr", True)
    If sDescr <> "" Then
        txtDescrTipo.Text = sDescr
    Else
        lbok = Mensaje("Ese Tipo No Existe", ICO_ERROR, False)
    End If
End If
End Sub

Private Sub txtTipo_LostFocus()
Dim sDescr As String
Dim lbok As Boolean

    sDescr = getDescrCatalogo(txtTipo, "Codigo", "vfacTipoFactura", "Descr", True)
    If sDescr <> "" Then
        txtDescrTipo.Text = sDescr
    Else
        lbok = Mensaje("Ese Tipo No Existe", ICO_ERROR, False)
    End If

End Sub



