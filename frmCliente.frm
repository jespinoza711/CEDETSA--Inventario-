VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCliente 
   Caption         =   "Form1"
   ClientHeight    =   10530
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15825
   LinkTopic       =   "Form1"
   ScaleHeight     =   10530
   ScaleWidth      =   15825
   StartUpPosition =   3  'Windows Default
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
      Left            =   4920
      Picture         =   "frmCliente.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Deshacer / Cancelar"
      Top             =   4470
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
      Left            =   4920
      Picture         =   "frmCliente.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Elimina el item actualmente seleccionado en el grid de datos ..."
      Top             =   3270
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
      Left            =   4920
      Picture         =   "frmCliente.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Agrega el item con los datos digitados..."
      Top             =   2070
      Width           =   555
   End
   Begin VB.CommandButton cmdSave 
      Enabled         =   0   'False
      Height          =   555
      Left            =   4920
      Picture         =   "frmCliente.frx":225E
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Aplica y Guarda los datos de la transacción en Firme ..."
      Top             =   3870
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
      Left            =   4920
      Picture         =   "frmCliente.frx":3F28
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Modifica los datos mostrados en el Grid con los datos digitados ..."
      Top             =   2670
      Width           =   555
   End
   Begin VB.TextBox txtCodigo 
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
      Height          =   285
      Left            =   6435
      TabIndex        =   2
      Top             =   1035
      Width           =   1095
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
      ForeColor       =   &H002F2F2F&
      Height          =   285
      Left            =   8835
      TabIndex        =   1
      Top             =   1035
      Width           =   7695
   End
   Begin VB.CheckBox chkCredito 
      Caption         =   "Cliente de Crédito ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   16560
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7575
      Left            =   180
      TabIndex        =   8
      Top             =   990
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   13361
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   3092271
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin TrueOleDBGrid60.TDBGrid TDBG 
      Height          =   1695
      Left            =   600
      OleObjectBlob   =   "frmCliente.frx":4BF2
      TabIndex        =   9
      Top             =   6600
      Visible         =   0   'False
      Width           =   3705
   End
   Begin ActiveTabs.SSActiveTabs SSActiveTabs1 
      Height          =   8535
      Left            =   6030
      TabIndex        =   13
      Top             =   1710
      Width           =   13410
      _ExtentX        =   23654
      _ExtentY        =   15055
      _Version        =   131083
      TabCount        =   3
      Tabs            =   "frmCliente.frx":A9AB
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   8145
         Left            =   30
         TabIndex        =   14
         Top             =   360
         Width           =   13350
         _ExtentX        =   23548
         _ExtentY        =   14367
         _Version        =   131083
         TabGuid         =   "frmCliente.frx":AA7C
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   8145
         Left            =   -99969
         TabIndex        =   15
         Top             =   360
         Width           =   13350
         _ExtentX        =   23548
         _ExtentY        =   14367
         _Version        =   131083
         TabGuid         =   "frmCliente.frx":AAA4
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   8145
         Index           =   0
         Left            =   30
         TabIndex        =   16
         Top             =   360
         Width           =   13350
         _ExtentX        =   23548
         _ExtentY        =   14367
         _Version        =   131083
         TabGuid         =   "frmCliente.frx":AACC
         Begin VB.Frame Frame2 
            Height          =   4215
            Left            =   240
            TabIndex        =   39
            Top             =   3720
            Width           =   13095
            Begin VB.CommandButton cmdDelclasif3 
               Height          =   320
               Left            =   3120
               Picture         =   "frmCliente.frx":AAF4
               Style           =   1  'Graphical
               TabIndex        =   73
               Top             =   1320
               Width           =   300
            End
            Begin VB.TextBox txtVendedor 
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
               Left            =   2160
               TabIndex        =   72
               Top             =   1320
               Width           =   855
            End
            Begin VB.TextBox txtDescrVendedor 
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
               Left            =   3600
               TabIndex        =   71
               Top             =   1320
               Width           =   8295
            End
            Begin VB.CommandButton cmdVendedor 
               Height          =   320
               Left            =   1560
               Picture         =   "frmCliente.frx":AF36
               Style           =   1  'Graphical
               TabIndex        =   70
               Top             =   1320
               Width           =   300
            End
            Begin VB.CommandButton cmdDelCategoria 
               Height          =   320
               Left            =   3120
               Picture         =   "frmCliente.frx":B278
               Style           =   1  'Graphical
               TabIndex        =   69
               Top             =   840
               Width           =   300
            End
            Begin VB.TextBox txtCategoria 
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
               Left            =   2160
               TabIndex        =   68
               Top             =   840
               Width           =   855
            End
            Begin VB.TextBox txtDescrCategoria 
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
               Left            =   3600
               TabIndex        =   67
               Top             =   840
               Width           =   8295
            End
            Begin VB.CommandButton cmdCategoria 
               Height          =   320
               Left            =   1560
               Picture         =   "frmCliente.frx":B6BA
               Style           =   1  'Graphical
               TabIndex        =   66
               Top             =   840
               Width           =   300
            End
            Begin VB.TextBox txtTecho 
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
               Left            =   10320
               TabIndex        =   65
               Top             =   1800
               Width           =   1215
            End
            Begin VB.CommandButton cmdDelBodeg 
               Height          =   320
               Left            =   3120
               Picture         =   "frmCliente.frx":B9FC
               Style           =   1  'Graphical
               TabIndex        =   64
               Top             =   360
               Width           =   300
            End
            Begin VB.TextBox txtBodega 
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
               Left            =   2160
               TabIndex        =   63
               Top             =   360
               Width           =   855
            End
            Begin VB.TextBox txtDescrBodega 
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
               Left            =   3600
               TabIndex        =   62
               Top             =   360
               Width           =   8295
            End
            Begin VB.CommandButton cmdBodega 
               Height          =   320
               Left            =   1560
               Picture         =   "frmCliente.frx":BE3E
               Style           =   1  'Graphical
               TabIndex        =   61
               Top             =   360
               Width           =   300
            End
            Begin VB.CommandButton cmdDelMoneda 
               Height          =   320
               Left            =   3120
               Picture         =   "frmCliente.frx":C180
               Style           =   1  'Graphical
               TabIndex        =   60
               Top             =   2280
               Width           =   300
            End
            Begin VB.TextBox txtMoneda 
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
               Left            =   2160
               TabIndex        =   59
               Top             =   2280
               Width           =   855
            End
            Begin VB.TextBox txtDescrMoneda 
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
               Left            =   3600
               TabIndex        =   58
               Top             =   2280
               Width           =   5055
            End
            Begin VB.CommandButton cmdMoneda 
               Height          =   320
               Left            =   1560
               Picture         =   "frmCliente.frx":C5C2
               Style           =   1  'Graphical
               TabIndex        =   57
               Top             =   2280
               Width           =   300
            End
            Begin VB.CommandButton cmdPlazo 
               Height          =   320
               Left            =   1560
               Picture         =   "frmCliente.frx":C904
               Style           =   1  'Graphical
               TabIndex        =   56
               Top             =   1800
               Width           =   300
            End
            Begin VB.TextBox txtDescrPlazo 
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
               Left            =   3600
               TabIndex        =   55
               Top             =   1800
               Width           =   5055
            End
            Begin VB.TextBox txtPlazo 
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
               Left            =   2160
               TabIndex        =   54
               Top             =   1800
               Width           =   855
            End
            Begin VB.CommandButton cmdDelPlazo 
               Height          =   320
               Left            =   3120
               Picture         =   "frmCliente.frx":CC46
               Style           =   1  'Graphical
               TabIndex        =   53
               Top             =   1800
               Width           =   300
            End
            Begin VB.TextBox txtSaldo 
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
               Left            =   10320
               TabIndex        =   52
               Top             =   2280
               Width           =   1215
            End
            Begin VB.CommandButton cmdDepto 
               Height          =   320
               Left            =   1560
               Picture         =   "frmCliente.frx":D088
               Style           =   1  'Graphical
               TabIndex        =   51
               Top             =   2760
               Width           =   300
            End
            Begin VB.TextBox txtDescrDepto 
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
               Left            =   3600
               TabIndex        =   50
               Top             =   2760
               Width           =   5055
            End
            Begin VB.TextBox txtDepartamento 
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
               Left            =   2160
               TabIndex        =   49
               Top             =   2760
               Width           =   855
            End
            Begin VB.CommandButton cmdDelDepto 
               Height          =   320
               Left            =   3120
               Picture         =   "frmCliente.frx":D3CA
               Style           =   1  'Graphical
               TabIndex        =   48
               Top             =   2760
               Width           =   300
            End
            Begin VB.CommandButton cmdMunicipio 
               Height          =   320
               Left            =   1560
               Picture         =   "frmCliente.frx":D80C
               Style           =   1  'Graphical
               TabIndex        =   47
               Top             =   3240
               Width           =   300
            End
            Begin VB.TextBox txtDescrMunicipio 
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
               Left            =   3600
               TabIndex        =   46
               Top             =   3240
               Width           =   5055
            End
            Begin VB.TextBox txtMunicipio 
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
               Left            =   2160
               TabIndex        =   45
               Top             =   3240
               Width           =   855
            End
            Begin VB.CommandButton cmdDelMuniicipio 
               Height          =   320
               Left            =   3120
               Picture         =   "frmCliente.frx":DB4E
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   3240
               Width           =   300
            End
            Begin VB.CommandButton cmdZona 
               Height          =   320
               Left            =   1560
               Picture         =   "frmCliente.frx":DF90
               Style           =   1  'Graphical
               TabIndex        =   43
               Top             =   3720
               Width           =   300
            End
            Begin VB.TextBox txtDescrZona 
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
               Left            =   3600
               TabIndex        =   42
               Top             =   3720
               Width           =   5055
            End
            Begin VB.TextBox txtZona 
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
               Left            =   2160
               TabIndex        =   41
               Top             =   3720
               Width           =   855
            End
            Begin VB.CommandButton cmdDelZona 
               Height          =   320
               Left            =   3120
               Picture         =   "frmCliente.frx":E2D2
               Style           =   1  'Graphical
               TabIndex        =   40
               Top             =   3720
               Width           =   300
            End
            Begin VB.Label lblVendedor 
               Caption         =   "Vendedor:"
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
               Left            =   240
               TabIndex        =   83
               Top             =   1320
               Width           =   1335
            End
            Begin VB.Label lblClasif 
               Caption         =   "Categoría:"
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
               Left            =   240
               TabIndex        =   82
               Top             =   840
               Width           =   1335
            End
            Begin VB.Label lblTecho 
               Caption         =   "Techo Crédito :"
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
               Left            =   8880
               TabIndex        =   81
               Top             =   1800
               Width           =   1335
            End
            Begin VB.Label lblClasif1 
               Caption         =   "Bodega Fact:"
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
               Left            =   240
               TabIndex        =   80
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label lblMoneda 
               Caption         =   "Moneda :"
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
               Left            =   240
               TabIndex        =   79
               Top             =   2280
               Width           =   1335
            End
            Begin VB.Label lblPlazo 
               Caption         =   "Plazo:"
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
               Left            =   240
               TabIndex        =   78
               Top             =   1800
               Width           =   1095
            End
            Begin VB.Label Label8 
               Caption         =   "Saldo Actual :"
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
               Left            =   8880
               TabIndex        =   77
               Top             =   2280
               Width           =   1335
            End
            Begin VB.Label Label9 
               Caption         =   "Departamento :"
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
               Left            =   240
               TabIndex        =   76
               Top             =   2760
               Width           =   1335
            End
            Begin VB.Label Label10 
               Caption         =   "Municipio:"
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
               Left            =   240
               TabIndex        =   75
               Top             =   3240
               Width           =   1335
            End
            Begin VB.Label Label11 
               Caption         =   "Zona:"
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
               Left            =   240
               TabIndex        =   74
               Top             =   3720
               Width           =   1335
            End
         End
         Begin VB.Frame Frame1 
            Height          =   3615
            Left            =   240
            TabIndex        =   17
            Top             =   120
            Width           =   12975
            Begin VB.TextBox txtRazon 
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
               Left            =   2040
               TabIndex        =   32
               Top             =   360
               Width           =   6975
            End
            Begin VB.CheckBox chkFarmacia 
               Caption         =   "Es Farmacia ?"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   9120
               TabIndex        =   31
               Top             =   360
               Width           =   1575
            End
            Begin VB.TextBox txtFarmacia 
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
               Left            =   2040
               TabIndex        =   30
               Top             =   840
               Width           =   6975
            End
            Begin VB.Frame Frame3 
               Caption         =   "Teléfonos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   120
               TabIndex        =   26
               Top             =   1200
               Width           =   6375
               Begin VB.TextBox txtTelefono1 
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
                  Left            =   240
                  TabIndex        =   29
                  Top             =   360
                  Width           =   1815
               End
               Begin VB.TextBox txtTelefono2 
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
                  Left            =   2400
                  TabIndex        =   28
                  Top             =   360
                  Width           =   1815
               End
               Begin VB.TextBox txtTelefono3 
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
                  Left            =   4440
                  TabIndex        =   27
                  Top             =   360
                  Width           =   1695
               End
            End
            Begin VB.Frame Frame4 
               Caption         =   "Celulares"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   6600
               TabIndex        =   23
               Top             =   1200
               Width           =   6135
               Begin VB.TextBox txtCelular2 
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
                  Left            =   3000
                  TabIndex        =   25
                  Top             =   360
                  Width           =   2175
               End
               Begin VB.TextBox txtCelular1 
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
                  Left            =   480
                  TabIndex        =   24
                  Top             =   360
                  Width           =   2055
               End
            End
            Begin VB.TextBox txtemail 
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
               Left            =   2160
               TabIndex        =   22
               Top             =   2280
               Width           =   9735
            End
            Begin VB.TextBox txtRUC 
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
               Left            =   9960
               TabIndex        =   21
               Top             =   840
               Width           =   2655
            End
            Begin VB.TextBox txtPropietario 
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
               Left            =   2160
               TabIndex        =   20
               Top             =   2760
               Width           =   9735
            End
            Begin VB.TextBox txtDireccion 
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
               Left            =   2160
               TabIndex        =   19
               Top             =   3240
               Width           =   9735
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
               Height          =   255
               Left            =   11520
               TabIndex        =   18
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label2 
               Caption         =   "Razón Social:"
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
               TabIndex        =   38
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label3 
               Caption         =   "Farmacia :"
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
               TabIndex        =   37
               Top             =   840
               Width           =   1335
            End
            Begin VB.Label Label5 
               Caption         =   "email:"
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
               Left            =   240
               TabIndex        =   36
               Top             =   2280
               Width           =   1335
            End
            Begin VB.Label Label6 
               Caption         =   "RUC:"
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
               Left            =   9120
               TabIndex        =   35
               Top             =   840
               Width           =   615
            End
            Begin VB.Label Label7 
               Caption         =   "Propietario :"
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
               Left            =   240
               TabIndex        =   34
               Top             =   2760
               Width           =   1335
            End
            Begin VB.Label Label12 
               Caption         =   "Dirección :"
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
               Left            =   240
               TabIndex        =   33
               Top             =   3240
               Width           =   1335
            End
         End
      End
   End
   Begin VB.Label lbFormCaption 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Catálogo de Vendedor"
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
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   19140
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   30
      Picture         =   "frmCliente.frx":E714
      Stretch         =   -1  'True
      Top             =   30
      Width           =   19590
   End
   Begin VB.Label Label1 
      Caption         =   "Código :"
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
      Left            =   5115
      TabIndex        =   12
      Top             =   1035
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Nombre:"
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
      Left            =   7635
      TabIndex        =   11
      Top             =   1035
      Width           =   1095
   End
End
Attribute VB_Name = "frmCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As ADODB.Recordset
Dim rst2 As ADODB.Recordset

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
End Sub

Public Sub HabilitarControles()
    Select Case Accion
        Case TypAccion.Add
            txtCodigo.Text = "1000"
            txtCodigo.Enabled = False
            txtDescr.Enabled = True
            chkActivo.Enabled = True
            chkActivo.value = 1
            chkFarmacia.value = 0
            chkCredito.Enabled = True
            chkCredito.value = 1
            
            txtDescr.Text = ""
            fmtTextbox txtCodigo, "R"
            fmtTextbox txtDescr, "O"
            
            txtRazon.Enabled = True
            txtRazon.Text = ""
            fmtTextbox txtRazon, "O"
            txtFarmacia.Enabled = True
            txtFarmacia.Text = ""
            txtRUC.Enabled = True
            txtRUC.Text = ""
            fmtTextbox txtRUC, "O"
            txtTelefono1.Enabled = True
            txtTelefono1.Text = ""
            txtTelefono2.Enabled = True
            txtTelefono2.Text = ""
            txtTelefono3.Enabled = True
            txtTelefono3.Text = ""
            txtCelular1.Enabled = True
            txtCelular1.Text = ""
            txtCelular2.Enabled = True
            txtCelular2.Text = ""
            txtemail.Enabled = True
            txtemail.Text = ""
            txtPropietario.Enabled = True
            txtPropietario.Text = ""
            txtTecho.Text = "0"
            txtTecho.Enabled = False
            txtSaldo.Text = "0"
            txtSaldo.Enabled = False
            
            txtBodega.Enabled = True
            txtBodega.Text = ""
            txtDescrBodega.Text = ""
            txtDescrBodega.Enabled = False
            txtCategoria.Enabled = True
            txtCategoria.Text = ""
            txtDescrCategoria.Text = ""
            txtDescrCategoria.Enabled = False
            txtVendedor.Enabled = True
            txtVendedor.Text = ""
            txtDescrVendedor.Text = ""
            txtDescrVendedor.Enabled = False
            txtPlazo.Enabled = True
            txtPlazo.Text = ""
            txtDescrPlazo.Text = ""
            txtDescrPlazo.Enabled = False
            txtMoneda.Enabled = True
            txtMoneda.Text = ""
            txtDescrMoneda.Text = ""
            txtDescrMoneda.Enabled = False
            txtDepartamento.Enabled = True
            txtDepartamento.Text = ""
            txtDescrDepto.Text = ""
            txtDescrDepto.Enabled = False
            txtMunicipio.Enabled = True
            txtMunicipio.Text = ""
            txtDescrMunicipio.Text = ""
            txtDescrMunicipio.Enabled = False
            txtZona.Enabled = True
            txtZona.Text = ""
            txtDescrZona.Text = ""
            txtDescrZona.Enabled = False
            TDBG.Enabled = False
            cmdBodega.Enabled = True
            cmdCategoria.Enabled = True
            cmdDepto.Enabled = True
            cmdMoneda.Enabled = True
            cmdMunicipio.Enabled = True
            cmdPlazo.Enabled = True
            cmdVendedor.Enabled = True
            cmdZona.Enabled = True
        Case TypAccion.Edit
            txtCodigo.Enabled = True
            txtDescr.Enabled = True
            fmtTextbox txtCodigo, "R"
            fmtTextbox txtDescr, "O"
            chkActivo.Enabled = True
            txtBodega.Enabled = True
            txtCategoria.Enabled = True
            txtVendedor.Enabled = True
            txtPlazo.Enabled = True
            txtMoneda.Enabled = True
            txtDepartamento.Enabled = False
            txtMunicipio.Enabled = False
            txtZona.Enabled = False
            txtDescrBodega.Enabled = False
            txtDescrCategoria.Enabled = False
            txtDescrVendedor.Enabled = False
            txtDescrPlazo.Enabled = False
            txtDescrMoneda.Enabled = False
            txtDescrDepto.Enabled = False
            txtDescrMunicipio.Enabled = False
            txtDescrZona.Enabled = False
            
            cmdBodega.Enabled = True
            cmdCategoria.Enabled = True
            cmdVendedor.Enabled = True
            cmdPlazo.Enabled = True
            cmdMoneda.Enabled = True
            cmdDepto.Enabled = True
            cmdMunicipio.Enabled = True
            cmdZona.Enabled = True
            
            TDBG.Enabled = False
        Case TypAccion.View
            txtCodigo.Enabled = False
            txtDescr.Enabled = False
            fmtTextbox txtCodigo, "R"
            fmtTextbox txtDescr, "O"
            chkActivo.Enabled = False
            txtBodega.Enabled = False
            txtCategoria.Enabled = False
            txtVendedor.Enabled = False
            txtPlazo.Enabled = False
            txtMoneda.Enabled = False
            txtDepartamento.Enabled = False
            txtMunicipio.Enabled = False
            txtZona.Enabled = False
            txtDescrBodega.Enabled = False
            txtDescrCategoria.Enabled = False
            txtDescrVendedor.Enabled = False
            txtDescrPlazo.Enabled = False
            txtDescrMoneda.Enabled = False
            txtDescrDepto.Enabled = False
            txtDescrMunicipio.Enabled = False
            txtDescrZona.Enabled = False
            
            cmdBodega.Enabled = False
            cmdCategoria.Enabled = False
            cmdVendedor.Enabled = False
            'TDBG.Enabled = False
    End Select
End Sub

Private Function CargaTablas() As Boolean
    Dim lbok As Boolean
    Dim iResultado As Integer
    On Error GoTo error
    lbok = True
      GSSQL = gsCompania & ".globalGetTablas -1 "
              
      'Set rst2 = gConet.Execute(GSSQL, adCmdText)  'Ejecuta la sentencia
      rst2.Open GSSQL, gConet, adOpenDynamic, adLockOptimistic
    
      If (gConet.Errors.Count > 0) Then  'Pregunta si hubo un error de ejecución
        gsOperacionError = "No existe ese cliente." 'Asigna msg de error
        lbok = False  'Indica que no es válido
        
      ElseIf Not (rst2.BOF And rst2.EOF) Then  'Si no es válido
        rst2.MoveNext

        lbok = True
      End If
      CargaTablas = lbok
      rst2.Close
      Exit Function
error:
      lbok = False
      gsOperacionError = "Ocurrió un error en la operación de carga de parametros " & err.Description
      Resume Next
End Function

Private Sub cmdAdd_Click()
    Accion = Add
    HabilitarBotones
    HabilitarControles
    txtDescr.SetFocus
End Sub

Private Sub cmdBodega_Click()
   Dim frm As frmBrowseCat

    Set frm = New frmBrowseCat
    frm.gsCaptionfrm = "Bodega" '& lblund.Caption
    frm.gsTablabrw = "invBodega"
    frm.gsCodigobrw = "IDBodega"
    frm.gbTypeCodeStr = False
    frm.gsDescrbrw = "Descr"
    frm.gbFiltra = False
    
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
      txtBodega.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      txtDescrBodega.Text = frm.gsDescrbrw
      fmtTextbox txtDescrBodega, "R"
    End If
End Sub

Private Sub cmdCategoria_Click()
   Dim frm As frmBrowseCat

    Set frm = New frmBrowseCat
    frm.gsCaptionfrm = "Bodega" '& lblund.Caption
    frm.gsTablabrw = "vglobalCategoriaCliente"
    frm.gsCodigobrw = "CODIGO"
    frm.gbTypeCodeStr = False
    frm.gsDescrbrw = "Descr"
    frm.gbFiltra = False
    
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
      txtCategoria.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      txtDescrCategoria.Text = frm.gsDescrbrw
      fmtTextbox txtDescrCategoria, "R"
    End If
End Sub

Private Sub cmdDepto_Click()
  Dim frm As frmBrowseCat

    Set frm = New frmBrowseCat
    frm.gsCaptionfrm = "Bodega" '& lblund.Caption
    frm.gsTablabrw = "vglobalDepartamento"
    frm.gsCodigobrw = "CODIGO"
    frm.gbTypeCodeStr = False
    frm.gsDescrbrw = "Descr"
    frm.gbFiltra = False
    
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
      txtDepartamento.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      txtDescrDepto.Text = frm.gsDescrbrw
      fmtTextbox txtDescrDepto, "R"
    End If
End Sub

Private Sub cmdEditItem_Click()
    Accion = Edit
    GetDataFromGridToControl
    HabilitarBotones
    HabilitarControles
End Sub

Private Sub cmdMoneda_Click()
  Dim frm As frmBrowseCat

    Set frm = New frmBrowseCat
    frm.gsCaptionfrm = "Bodega" '& lblund.Caption
    frm.gsTablabrw = "vglobalMONEDA"
    frm.gsCodigobrw = "CODIGO"
    frm.gbTypeCodeStr = False
    frm.gsDescrbrw = "Descr"
    frm.gbFiltra = False
    
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
      txtMoneda.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      txtDescrMoneda.Text = frm.gsDescrbrw
      fmtTextbox txtDescrMoneda, "R"
    End If
End Sub

Private Sub cmdMunicipio_Click()
  Dim frm As frmBrowseCat

    Set frm = New frmBrowseCat
    frm.gsCaptionfrm = "Bodega" '& lblund.Caption
    frm.gsTablabrw = "vglobalMunicipio"
    frm.gsCodigobrw = "CODIGO"
    frm.gbTypeCodeStr = False
    frm.gsDescrbrw = "Descr"
    frm.gbFiltra = False
    
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
      txtMunicipio.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      txtDescrMunicipio.Text = frm.gsDescrbrw
      fmtTextbox txtDescrMunicipio, "R"
    End If
End Sub

Private Sub cmdPlazo_Click()
   Dim frm As frmBrowseCat

    Set frm = New frmBrowseCat
    frm.gsCaptionfrm = "Bodega" '& lblund.Caption
    frm.gsTablabrw = "vglobalPLAZO"
    frm.gsCodigobrw = "CODIGO"
    frm.gbTypeCodeStr = False
    frm.gsDescrbrw = "Descr"
    frm.gbFiltra = False
    
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
      txtPlazo.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      txtDescrPlazo.Text = frm.gsDescrbrw
      fmtTextbox txtDescrPlazo, "R"
    End If
End Sub

Private Sub cmdSave_Click()
    Dim lbok As Boolean
    Dim sMsg As String
    Dim sActivo As String
    Dim sFiltro As String
    
    Dim sEsFarmacia As String
    
        If Not ControlsOk() Then
            Exit Sub
        End If
        
        If chkActivo.value = 1 Then
            sActivo = "1"
        Else
            sActivo = "0"
        End If
        
         
        If chkFarmacia.value = 1 Then
            sEsFarmacia = "1"
        Else
            sEsFarmacia = "0"
        End If
        
            
    If Accion = Add Then
    
        If Not (rst.EOF And rst.BOF) Then
            sFiltro = "IDCliente = " & txtCodigo.Text
            If ExiteRstKey(rst, sFiltro) Then
               lbok = Mensaje("Ya existe ese Cliente ", ICO_ERROR, False)
                txtCodigo.SetFocus
            Exit Sub
            End If
        End If
    
                lbok = ccUpdateCliente("I", txtCodigo.Text, txtDescr.Text, txtRazon.Text, txtDireccion.Text, txtTelefono1.Text, txtTelefono2.Text, txtTelefono3.Text, txtCelular1.Text, txtCelular2.Text, _
                    txtemail.Text, sEsFarmacia, txtFarmacia.Text, txtRUC.Text, txtPropietario.Text, txtBodega.Text, txtPlazo.Text, txtMoneda.Text, txtCategoria.Text, txtDepartamento.Text, _
                    txtMunicipio.Text, txtZona.Text, txtVendedor.Text, txtTecho.Text, sActivo, gsUSUARIO)
            
            If lbok Then
                sMsg = "El Cliente ha sido registrado exitosamente ... "
                lbok = Mensaje(sMsg, ICO_OK, False)
                ' actualiza datos
                cargaGrid
                Accion = View
                HabilitarControles
                HabilitarBotones
            Else
                 sMsg = "Ha ocurrido un error tratando de Agregar el Cliente ... "
                lbok = Mensaje(sMsg, ICO_ERROR, False)
            End If
    End If ' si estoy adicionando
    If Accion = Edit Then
        If Not (rst.EOF And rst.BOF) Then
                lbok = ccUpdateCliente("U", txtCodigo.Text, txtDescr.Text, txtRazon.Text, txtDireccion.Text, txtTelefono1.Text, txtTelefono2.Text, txtTelefono3.Text, txtCelular1.Text, txtCelular2.Text, _
                    txtemail.Text, sEsFarmacia, txtFarmacia.Text, txtRUC.Text, txtPropietario.Text, txtBodega.Text, txtPlazo.Text, txtMoneda.Text, txtCategoria.Text, txtDepartamento.Text, _
                    txtMunicipio.Text, txtZona.Text, txtVendedor.Text, txtTecho.Text, sActivo, gsUSUARIO)
    
            If lbok Then
                sMsg = "Los datos fueron grabados Exitosamente ... "
                lbok = Mensaje(sMsg, ICO_OK, False)
                ' actualiza datos
                cargaGrid
                Accion = View
                HabilitarControles
                HabilitarBotones
            Else
                sMsg = "Ha ocurrido un error tratando de actualizar los datos del producto... "
                lbok = Mensaje(sMsg, ICO_ERROR, False)
            End If
        End If
       End If ' si estoy adicionando

End Sub

Private Sub cmdVendedor_Click()
   Dim frm As frmBrowseCat

    Set frm = New frmBrowseCat
    frm.gsCaptionfrm = "Vendedor" '& lblund.Caption
    frm.gsTablabrw = "fafVendedor"
    frm.gsCodigobrw = "IDVendedor"
    frm.gbTypeCodeStr = False
    frm.gsDescrbrw = "Nombre"
    frm.gbFiltra = False
    
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
      txtVendedor.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      txtDescrVendedor.Text = frm.gsDescrbrw
      fmtTextbox txtDescrVendedor, "R"
    End If
End Sub

Private Sub cmdZona_Click()
  Dim frm As frmBrowseCat

    Set frm = New frmBrowseCat
    frm.gsCaptionfrm = "Bodega" '& lblund.Caption
    frm.gsTablabrw = "vglobalZona"
    frm.gsCodigobrw = "CODIGO"
    frm.gbTypeCodeStr = False
    frm.gsDescrbrw = "Descr"
    frm.gbFiltra = False
    
    frm.Show vbModal
    If frm.gsCodigobrw <> "" Then
      txtZona.Text = frm.gsCodigobrw
      
    End If
    
    If frm.gsDescrbrw <> "" Then
      txtDescrZona.Text = frm.gsDescrbrw
      fmtTextbox txtDescrZona, "R"
    End If
End Sub

Private Sub Form_Load()
    Dim lbok As Boolean
    Set rst = New ADODB.Recordset
    If rst.State = adStateOpen Then rst.Close
    rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
    rst.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
    rst.CursorLocation = adUseClient ' Cursor local al cliente
    rst.LockType = adLockOptimistic
    
    Set rst2 = New ADODB.Recordset
    If rst2.State = adStateOpen Then rst2.Close
    rst2.ActiveConnection = gConet 'Asocia la conexión de trabajo
    rst2.CursorType = adOpenStatic 'adOpenKeyset  'Asigna un cursor dinamico
    rst2.CursorLocation = adUseClient ' Cursor local al cliente
    rst2.LockType = adLockOptimistic
    
    
    Caption = gsFormCaption
    lbFormCaption = gsTitle
    
    Accion = View
    HabilitarControles
    HabilitarBotones
    lbok = CargaTablas()
    cargaGrid
End Sub
Private Sub IniciaIconos()
cmdSave.Enabled = False
cmdEditItem.Enabled = True
cmdEliminar.Enabled = True
cmdAdd.Enabled = True
bEdit = False
bAdd = False

End Sub
Private Sub InicializaListView()
    Dim i As Long
    Dim sItem As String
        With ListView1
            ' Las pruebas serán en modo "detalle"
            .View = lvwReport
            ' al seleccionar un elemento, seleccionar la línea completa
            .FullRowSelect = True
            ' Mostrar las líneas de la cuadrícula
            .GridLines = True
            ' No permitir la edición automática del texto
            .LabelEdit = lvwManual
            ' Permitir múltiple selección
            .MultiSelect = False
            ' Para que al perder el foco,
            ' se siga viendo el que está seleccionado
            .HideSelection = False
            .LabelWrap = False
            .ForeColor = vbBlue
    
        End With
    
        With ListView1.ColumnHeaders.Add(, , "Descripción", 4500)
            
                .Tag = cTexto
        End With
    
     
        '
        ' Eliminar las cabeceras
        ListView1.ColumnHeaders.Clear
        '
        ' Asignar las cabeceras
        With ListView1.ColumnHeaders.Add(, , "***************   CLIENTES   **************", 4500)
            .Tag = cTexto
        End With
        
     ListView1.ListItems.Clear
        ' Asignar algunos datos aleatorios
        If Not (rst.EOF And rst.BOF) Then
            rst.MoveFirst
            While Not rst.EOF
            
            sItem = Trim(right("00000" + Trim(Str(rst("IDCliente").value)), 5)) + "-" + rst("NOMBRE").value
            If Len(sItem) > 50 Then
                sItem = Mid(sItem, 1, 50) & vbLf & Mid(sItem, 51, Len(sItem))
                
            End If
            
                With ListView1.ListItems.Add(, , sItem)
                    ' Cada subitem debe corresponder con cada una de las cabeceras
                    ' la segunda cabecera es el Subitems(1) y así sucesivamente
                End With
                rst.MoveNext
            Wend
        End If

End Sub

Private Sub cargaGrid()
    Dim sIndependiente As String
    If rst.State = adStateOpen Then rst.Close
    rst.ActiveConnection = gConet 'Asocia la conexión de trabajo
    rst.CursorType = adOpenKeyset 'adOpenKeyset  'Asigna un cursor dinamico
    rst.CursorLocation = adUseClient ' Cursor local al cliente
    rst.LockType = adLockOptimistic
    GSSQL = gsCompania & ".fafgetClientes -1"
    If rst.State = adStateOpen Then rst.Close
    Set rst = GetRecordset(GSSQL)
    If Not (rst.EOF And rst.BOF) Then
      Set TDBG.DataSource = rst
      'CargarDatos rst, TDBG, "Codigo", "Descr"
      TDBG.Refresh
      'IniciaIconos
      rst.MoveFirst
                InicializaListView
                ListView1.ListItems(1).Selected = True
                ShowSelectedItem
                IniciaIconos
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (rst Is Nothing) Then Set rst = Nothing
    If Not (rst2 Is Nothing) Then Set rst2 = Nothing
End Sub

Private Sub ListView1_Click()
    ShowSelectedItem
    
End Sub


Private Sub ShowSelectedItem()
    Dim sValor As String
    Dim sFiltro As String
    Dim i As Integer
    If Not (rst.EOF And rst.BOF) Then
        With ListView1.SelectedItem
            sValor = .Text
            .ToolTipText = sValor
            i = InStr(sValor, "-")
            i = i - 1
            sValor = Mid(sValor, 1, i)
            sFiltro = "IDCliente=" & Str(Val(sValor))
            rst.MoveFirst
            rst.Find sFiltro
            'rst.Bookmark = getPositionRecord(rst, sFiltro)
            If Not (rst.EOF And rst.BOF) Then
                GetDataFromGridToControl
            End If
            
        End With
    
    End If
End Sub


Private Sub SSActiveTabs1_BeforeTabClick(ByVal NewTab As ActiveTabs.SSTab, ByVal Cancel As ActiveTabs.SSReturnBoolean)
    If NewTab.Key = "GENERAL" Then
    '    SSActiveTabs1.Height = 3255
        'TDBGMov.Height = 3255
    End If
    If NewTab.Key = "PRECIOS" Then
    '    SSActiveTabs1.Height = 3255
        'TDBGMov.Height = 3255
    End If
    
    If NewTab.Key = "EXISTENCIA" Then
        txtBodega.Text = ""
        txtDescrBodega.Text = ""
    '    SSActiveTabs1.Height = 3255
        'TDBGMov.Height = 3255
    End If
    If NewTab.Key = "MOVIMIENTOS" Then
    '    SSActiveTabs1.Height = Me.Height - 200
        'TDBGMov.Height = Me.Height - 100
    End If

End Sub

Private Sub TDBG_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    GetDataFromGridToControl
End Sub

Private Sub GetDataFromGridToControl()
    If Not (rst.EOF And rst.BOF) Then
        txtCodigo.Text = rst("IDCLIENTE").value
        txtDescr.Text = rst("NOMBRE").value
        txtRazon.Text = rst("Direccion").value
        txtTelefono1.Text = rst("Telefono1").value
        txtTelefono2.Text = rst("Telefono2").value
        txtTelefono3.Text = rst("Telefono3").value
        txtCelular1.Text = rst("Celular1").value
        txtCelular2.Text = rst("Celular2").value
        txtemail.Text = rst("email").value
        
        If rst("Activo").value = True Then
            chkActivo.value = 1
        Else
            chkActivo.value = 0
        End If
        
        If rst("EsFarmacia").value = True Then
            chkFarmacia.value = 1
        Else
            chkFarmacia.value = 0
        End If
        If rst("EsFarmacia").value = True Then
            chkFarmacia.value = 1
        Else
            chkFarmacia.value = 0
        End If
        txtFarmacia.Text = rst("NombreFarmacia").value
        txtRUC.Text = rst("RUC").value
        txtPropietario.Text = rst("Propietario").value
        txtBodega.Text = rst("IDBodega").value
        txtDescrBodega.Text = rst("DESCRBodega").value
        txtPlazo.Text = rst("IDPlazo").value
        txtDescrPlazo.Text = rst("DescrPlazo").value
        txtMoneda.Text = rst("IDMoneda").value
        txtDescrMoneda.Text = rst("DescrMoneda").value
        
        txtCategoria.Text = rst("IDCategoria").value
        txtDescrCategoria.Text = rst("DescrCatCliente").value
        
        txtDepartamento.Text = rst("IDDepartamento").value
        txtDescrDepto.Text = rst("DescrDepto").value
        
        txtMunicipio.Text = rst("IDMunicipio").value
        txtDescrMunicipio.Text = rst("DescrMunicipio").value
        txtZona.Text = rst("IDZona").value
        txtDescrZona.Text = rst("DescrZona").value
        
        txtVendedor.Text = rst("IDVendedor").value
        txtDescrVendedor.Text = rst("NombreVendedor").value
        txtSaldo.Text = rst("SaldoLocal").value
        txtTecho.Text = rst("TechoCredito").value
        
        
    Else
        txtCodigo.Text = ""
        txtDescr.Text = ""
    End If

End Sub

Private Function ControlsOk() As Boolean
    Dim lbok As Boolean
    If txtCodigo.Text = "" Then
        lbok = Mensaje("El Código del Producto no puede quedar en Blanco...", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    If Not Val_TextboxNum(txtCodigo) Then
        lbok = Mensaje("El Código del Producto debe ser Numérico", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    
    
    If txtDescr.Text = "" Then
        lbok = Mensaje("La Descripción del Producto no puede quedar en Blanco...", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    If txtRazon.Text = "" Then
        lbok = Mensaje("La Razón Social no puede quedar en Blanco...", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    If chkFarmacia.value = 1 And txtFarmacia.Text = "" Then
        lbok = Mensaje("El nombre de la farmacia no puede quedar en Blanco, cuando el cliente es Farmacia", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    If txtRUC.Text = "" Then
        lbok = Mensaje("El RUC no puede quedar en Blanco...", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    
    If txtBodega.Text = "" Then
        lbok = Mensaje("La Bodega de Facturación del Cliente no puede quedar en Blanco...", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    If txtCategoria.Text = "" Then
        lbok = Mensaje("La Categoría del Cliente no puede quedar en Blanco...", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    If txtVendedor.Text = "" Then
        lbok = Mensaje("El Vendedor que corresponde al Cliente no puede quedar en Blanco...", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    If txtPlazo.Text = "" Then
        lbok = Mensaje("El Plazo que corresponde al Cliente no puede quedar en Blanco...", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    If txtMoneda.Text = "" Then
        lbok = Mensaje("La Moneda del Cliente no puede quedar en Blanco...", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    If txtDepartamento.Text = "" Then
        lbok = Mensaje("El Departamento del Cliente no puede quedar en Blanco...", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    If txtMunicipio.Text = "" Then
        lbok = Mensaje("El Municipio del Cliente no puede quedar en Blanco...", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    If txtZona.Text = "" Then
        lbok = Mensaje("La Zona del Cliente no puede quedar en Blanco...", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    
    If txtTecho.Text = "" Then
        lbok = Mensaje("El Techo de Crédito del Cliente no puede quedar en Blanco...", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    
    If Not Val_TextboxNum(txtTecho) Then
        lbok = Mensaje("El Costo Ultimo Dolar del Producto debe ser Numérico", ICO_ERROR, False)
        ControlsOk = False
        Exit Function
    End If
    

    
    ControlsOk = True
End Function




