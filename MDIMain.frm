VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   10905
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   16080
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar statusMain 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   8
      Top             =   10620
      Width           =   16080
      _ExtentX        =   28363
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   11
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Picture         =   "MDIMain.frx":0000
            Text            =   "Usuario:"
            TextSave        =   "Usuario:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "UserName"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Picture         =   "MDIMain.frx":059A
            Key             =   "DataBase"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Picture         =   "MDIMain.frx":0B34
            Key             =   "Server"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10372
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
            Text            =   "Fecha:"
            TextSave        =   "Fecha:"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "22/05/2014"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "16:19"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NÚM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   16080
      TabIndex        =   7
      Top             =   10245
      Width           =   16080
      Begin VB.ListBox WinList 
         Height          =   255
         ItemData        =   "MDIMain.frx":10CE
         Left            =   270
         List            =   "MDIMain.frx":10D0
         TabIndex        =   9
         Top             =   30
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Timer tmrResize 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7890
      Top             =   4470
   End
   Begin VB.PictureBox picLeft 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   9480
      Left            =   0
      ScaleHeight     =   9480
      ScaleWidth      =   2310
      TabIndex        =   3
      Top             =   765
      Width           =   2310
      Begin VB.Frame Frame1 
         Height          =   465
         Left            =   0
         TabIndex        =   4
         Top             =   -75
         Width           =   2250
         Begin VB.Image Image 
            Height          =   240
            Index           =   0
            Left            =   75
            Picture         =   "MDIMain.frx":10D2
            Top             =   150
            Width           =   240
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   375
            TabIndex        =   5
            Top             =   195
            Width           =   1290
         End
      End
      Begin MSComctlLib.ListView lvWin 
         Height          =   4050
         Left            =   0
         TabIndex        =   6
         Top             =   390
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   7144
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "MDIMain.frx":1AD4
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Form Name"
            Object.Width           =   3969
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Image Image5 
         Height          =   960
         Left            =   1650
         Picture         =   "MDIMain.frx":27AE
         Top             =   4650
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   960
         Left            =   1650
         Picture         =   "MDIMain.frx":34F8
         Top             =   5730
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.PictureBox picSeparator 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   9480
      Left            =   2310
      MousePointer    =   9  'Size W E
      ScaleHeight     =   9480
      ScaleWidth      =   120
      TabIndex        =   2
      Top             =   765
      Width           =   120
      Begin VB.PictureBox StyleButton2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   -30
         ScaleHeight     =   1575
         ScaleWidth      =   60
         TabIndex        =   10
         Top             =   2790
         Width           =   60
      End
   End
   Begin VB.PictureBox picContainer 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   16080
      TabIndex        =   0
      Top             =   0
      Width           =   16080
      Begin MSComctlLib.Toolbar tbMenu 
         Height          =   780
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   14970
         _ExtentX        =   26405
         _ExtentY        =   1376
         ButtonWidth     =   2143
         ButtonHeight    =   1376
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "itb32x"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   23
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Izquierda"
               Key             =   "Izquierda"
               ImageIndex      =   23
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Derecha"
               Key             =   "Derecha"
               ImageIndex      =   24
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   14
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ajustar"
               Key             =   "Ajustar"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Chequear Todo"
               Key             =   "Chequear Todo"
               ImageIndex      =   20
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Des Chequear"
               Key             =   "Des Chequear"
               ImageIndex      =   21
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Filtro"
               Key             =   "Filtro"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Nuevo"
               Key             =   "Nuevo"
               Object.ToolTipText     =   "Ctrl+F2"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ver"
               Key             =   "Ver"
               ImageIndex      =   18
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Editar"
               Key             =   "Editar"
               Object.ToolTipText     =   "Ctrl+F3"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar"
               Key             =   "Eliminar"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Guardar"
               Key             =   "Guardar"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cancelar"
               Key             =   "Cancelar"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Aplicar"
               Key             =   "Aplicar"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Buscar"
               Key             =   "Buscar"
               Object.ToolTipText     =   "Ctrl+F4"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Limpiar"
               Key             =   "Limpiar"
               ImageIndex      =   22
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Refrescar"
               Key             =   "Refrescar"
               Object.ToolTipText     =   "Ctrl+F6"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Imprimir"
               Key             =   "Imprimir"
               Object.ToolTipText     =   "Ctrl+F7"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Exportar"
               Key             =   "Exportar"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cerrar"
               Key             =   "Cerrar"
               Object.ToolTipText     =   "Ctrl+F8"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "User's Guide"
               Key             =   "User's Guide"
               ImageIndex      =   15
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Acerca de"
               Key             =   "Acerca"
               ImageIndex      =   16
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cerrar Todo"
               Key             =   "CerrarTodo"
               ImageIndex      =   17
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5850
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":4242
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":5BD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":68B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":8242
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":9BD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":B566
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":CEF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":DBD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":E8AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":F586
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":10262
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":10F3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":1181A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":124F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":131D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":13EAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":14792
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":1546E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":15D4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":16A26
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":183BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":19D4E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList itb32x 
      Left            =   5760
      Top             =   3420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":1A62A
            Key             =   "NEW"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":1BFBC
            Key             =   "EDIT"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":1D94E
            Key             =   "SEARCH"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":1F2E0
            Key             =   "DELETE"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":20C72
            Key             =   "save"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":2294C
            Key             =   "REFRESH"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":242DE
            Key             =   "PRINT"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":25C70
            Key             =   "UNDO"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":2694A
            Key             =   "FILTER"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":26EE4
            Key             =   "EXPORT EXCEL"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":27B5F
            Key             =   "APLICAR"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":28839
            Key             =   "CLOSE"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":2A1CB
            Key             =   "SHORTCUTS"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":2BB5D
            Key             =   "ADJUST"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":2D4F1
            Key             =   "USERSGUIDE"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":2E1CD
            Key             =   "ABOUT"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":2EAAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":2F789
            Key             =   "VIEW"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":30465
            Key             =   "VOID"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":31141
            Key             =   "CHECKALL"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":31E1D
            Key             =   "UNCHECKALL"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":32AF9
            Key             =   "CLEAR"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":333D5
            Key             =   "LEFT"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":33827
            Key             =   "RIGHT"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "Edicion"
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
   End
   Begin VB.Menu mnuAdministracion 
      Caption         =   "Administracion"
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cursor_pos As POINTAPI
Dim resize_down     As Boolean
Dim show_mnu        As Boolean
Dim pos_num         As Integer
Dim WindowsOpen     As Integer


Public Sub AddForm(frmName As String)
'Add form name to list
'This sub is called from every form outside
    Dim i As Integer
    i = IsWindowInListbox(frmName, WinList)
    If i = -1 Then
        WinList.AddItem (frmName)
        WindowsOpen = WinList.ListCount
    End If

End Sub

Public Sub SubtractForm(frmName As String)
'This sub is called from every form outside
    Dim i As Integer
    If WinList.ListCount < 1 Then Exit Sub
    i = IsWindowInListbox(frmName, WinList)
    If i > -1 Then
        WinList.RemoveItem (i)
        WindowsOpen = WinList.ListCount
    End If
    
End Sub




Private Sub lvWin_Click()
  Dim s As String
    Dim i As Integer
    
    'If lvWin.ListItems.Count < 1 Then Exit Sub
    

    
    
    Select Case lvWin.SelectedItem.Key
        'Case "frmShortcuts": frmShortcuts.show: frmShortcuts.WindowState = vbMaximized: frmShortcuts.SetFocus
    
        'Inventory
        Case "frmProductos":
            Dim oformProductos As New frmProductos
            oformProductos.gsFormCaption = "Maestro de Productos"
            oformProductos.gsTitle = "MAESTRO DE PRODUCTOS"
            LoadForm oformProductos
        Case "frmMasterLotes":
            Dim ofrmLotes  As New frmMasterLotes
            ofrmLotes.gsFormCaption = "Maestro de Lotes"
            ofrmLotes.gsTitle = "MAESTRO DE LOTES DE PRODUCTOS"
            LoadForm ofrmLotes
        Case "frmBodega":
            Dim ofrmBodega  As New frmBodega
            ofrmBodega.gsFormCaption = "Catálogo de Bodegas"
            ofrmBodega.gsTitle = "BODEGAS"
            LoadForm ofrmBodega
        Case "frmTransacciones":
            Dim ofrmTran As New frmTransacciones
            ofrmTran.gsFormCaption = "Transacciones"
            ofrmTran.gsTitle = "TRANSACCIONES DE PRODUCTOS"
            LoadForm ofrmTran
        Case "frmVendedor":
            Dim ofrmVendedor As New frmVendedor
            ofrmVendedor.gsFormCaption = "Catalogo de Vendedores"
            ofrmVendedor.gsTitle = "VENDEDORES"
            LoadForm ofrmVendedor
        Case "frmCliente":
            Dim ofrmCliente As New frmCliente
            ofrmCliente.gsTitle = "Clientes"
            ofrmCliente.gsTitle = "MAESTRO DE CLIENTES"
            LoadForm ofrmCliente
        Case "frmPedidos":
            Dim ofrmPedidos As New frmPedidos
            ofrmPedidos.gsFormCaption = "Pedidos"
            ofrmPedidos.gsTitle = "TELE VENTAS"
            LoadForm ofrmPedidos
        Case "frmTablas":
            Dim ofrmTablas As New frmTablas
            ofrmTablas.gsFormCaption = "Catalogo de Tablas Generales"
            ofrmTablas.gsTitle = "TABLAS GENERALES DEL SISTEMA"
            LoadForm ofrmTablas
        Case "frmCatalogos":
            Dim ofrmCatalogos As New frmCatalogos
            ofrmCatalogos.gsFormCaption = "Catalogo Generales"
            ofrmCatalogos.gsTitle = "CATALOGOS GENERALES DEL SISTEMA"
            LoadForm ofrmCatalogos
       '----------------------------------------------
                     
        
    End Select
End Sub

Private Sub MostrarDatosUsuario()
    Me.statusMain.Panels(2).Text = gsUser
    Me.statusMain.Panels(3).Text = "DataBase: " & gsNombreBaseDatos
    Me.statusMain.Panels(4).Text = "Server: " & gsNombreServidor
End Sub


Private Sub MDIForm_Load()
    WindowsOpen = 0
    WinList.Clear
    Call SetupMenuButtons                               'Initialise the buttons array
    Call SetupFormToolbar("no form")
    MostrarDatosUsuario
    Me.Show
     Set lvWin.SmallIcons = ImageList1
    Set lvWin.Icons = ImageList1
     Call lvWin_Load
     Call MagicCusror(200)
     
     show_mnu = True
    show_menu (show_mnu)
End Sub


Private Sub show_menu(ByVal Show As Boolean)
    Dim img As Image
    If Show = True Then
        Set img = Image2
    Else
        Set img = Image5
    End If
    'Set the style button graphics
    With StyleButton2
        Set .Picture = img.Picture
'        Set .PictureFocus = img.Picture
'        Set .PictureHover = img.Picture
'        Set .PictureUp = img.Picture
    End With
    'Set picture visibility
    picLeft.Visible = Show
    
    If Show = True Then StyleButton2.ToolTipText = "Hide": picSeparator.MousePointer = vbSizeWE Else picSeparator.MousePointer = vbArrow: StyleButton2.ToolTipText = "Expand"
    
    Set img = Nothing
End Sub



Private Sub lvWin_Load()
'Michael's sub
    
    
    With lvWin
        .ListItems.Clear

        Set .SmallIcons = ImageList1
        Set .Icons = ImageList1
        'For Sales
       
        .ListItems.Add(, "frmProductos", "Maestro Productos", 6, 6).Bold = False
        .ListItems.Add(, "frmMasterLotes", "Maestro de Lotes", 5, 5).Bold = False
        .ListItems.Add(, "frmBodega", "Bodegas", 21, 21).Bold = False
        
        .ListItems.Add(, "frmTransacciones", "Transacciones Producto", 16, 16).Bold = False
        
        .ListItems.Add(, "frmVendedor", "Vendedores", 3, 3).Bold = False
        .ListItems.Add(, "frmCliente", "Clientes", 4, 4).Bold = False
        

        .ListItems.Add(, "frmPedidos", "Pedidos", 12, 12).Bold = False
        .ListItems.Add(, "frmTablas", "Tablas Globales", 13, 13).Bold = False
'
'        'For Inventory
        .ListItems.Add(, "frmCatalogos", "Catalogos Generales", 4, 4).Bold = False
'
'        .ListItems.Add(, "frmCategories", "Category List", 5, 5).Bold = False
'        .ListItems.Add(, "frmProduct", "Product List", 6, 6).Bold = False
'
'        .ListItems.Add(, "frmStockMonitoring", "Stock Monitoring", 9, 9).Bold = False
'        .ListItems.Add(, "frmStockReceive", "Stock Receive", 8, 8).Bold = False
'
'        'For Transaction
'        .ListItems.Add(, "frmLoading", "Van Loading", 10, 10).Bold = False
'        .ListItems.Add(, "frmInvoice", "Sales Invoice", 14, 14).Bold = False
'        .ListItems.Add(, "frmVanCollection", "Van Collection", 15, 15).Bold = False
'        .ListItems.Add(, "frmVanInventory", "Van Inventory", 11, 11).Bold = False
'        .ListItems.Add(, "frmVanRemmitance", "Remmitance", 19, 19).Bold = False
'
'        .ListItems.Add(, "frmSelectZipCode", "Manage Zip Codes", 20, 20).Bold = False
'        .ListItems.Add(, "frmSelectBank", "Manage Bank Records", 21, 21).Bold = False
'        .ListItems.Add(, "frmUserRec", "User Records", 17, 17).Bold = False
'        .ListItems.Add(, "frmBusinessInfo", "Business Information", 16, 16).Bold = False
    End With
End Sub

Sub MagicCusror(x As Integer)
 picLeft.Width = picLeft + (x * Screen.TwipsPerPixelX) - (Me.left + 110)

End Sub



Private Sub picSeparator_Resize()
    Call center_obj_vertical(picSeparator, StyleButton2)
End Sub

Private Sub picLeft_Resize()
    On Error Resume Next
    Frame1.Width = picLeft.ScaleWidth
    lvWin.Width = picLeft.ScaleWidth
    lvWin.Height = picLeft.ScaleHeight - lvWin.top - 20
End Sub

Private Sub picSeparator_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If show_mnu = False Then Exit Sub
    If Button = vbLeftButton Then
        tmrResize.Enabled = True
        resize_down = True
    End If
End Sub

Private Sub picSeparator_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If show_mnu = False Then Exit Sub
    If Button = vbLeftButton Then
        tmrResize.Enabled = False
        resize_down = False
    End If
End Sub



Private Sub StyleButton2_Click()
   
        show_mnu = Not show_mnu
        show_menu show_mnu

End Sub

Private Sub tmrResize_Timer()
    On Error Resume Next
    GetCursorPos cursor_pos
    'picLeft.Width = (Me.Width - ((cursor_pos.x * Screen.TwipsPerPixelX) - Me.Left)) - 90
   
    picLeft.Width = picLeft + (cursor_pos.x * Screen.TwipsPerPixelX) - (Me.left + 110)
    
End Sub



Public Sub UnloadChilds()
''Unload all active forms
Dim Form As Form
   For Each Form In Forms
      ''Unload all active childs
      If Form.Name <> Me.Name And Form.Name <> "frmShortcuts" Then Unload Form
   Next Form
   
Set Form = Nothing
End Sub



Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)

    
    Select Case Button.Key
        Case "Izquierda":
            Call GoLeft
            
        Case "Derecha":
            Call GoRight
            
        Case "Acerca":
            Call ShowAbout
            
        Case "User's Guide":
            'Call mnuHUG_Click
            
        Case "CerrarTodo":
            Call CloseAll
        
        
        Case Else:
            On Error Resume Next
           ActiveForm.CommandPass Button.Key
    End Select
    
   
End Sub

Public Sub GoLeft()
'try to activate a previous child window

  Dim i As Integer
  Dim frm As String
  Dim itmFound As ListItem   ' FoundItem variable.
  
  If WindowsOpen = 0 Then Exit Sub
  frm = ActiveForm.Name
  
    i = IsWindowInListbox(frm, WinList)
    If i < 1 Then Exit Sub
    frm = WinList.List(i - 1)
    'If itmFound Is Nothing Then Exit Sub
    
   Set itmFound = lvWin.ListItems(frm)  'frm is the key to the obj . itmFound is an object of type ListItem
    Set lvWin.SelectedItem = itmFound   'select this obj in the lvWin control array

   Call lvWin_Click                     'pretend the user clicked in the listview to selected this form window.
   
   
   

End Sub

Public Sub GoRight()
'try to activate a next child window that should already be open

  Dim i As Integer
  Dim frm As String
  Dim itmFound As ListItem                  ' FoundItem variable.
  
  If WindowsOpen < 2 Then Exit Sub
  frm = ActiveForm.Name
  
    i = IsWindowInListbox(frm, WinList)
    If i = -1 Then Exit Sub
    If i = WindowsOpen - 1 Then Exit Sub
    frm = WinList.List(i + 1)
    
    
   Set itmFound = lvWin.ListItems(frm)      'frm is the key to the obj.
    Set lvWin.SelectedItem = itmFound

   Call lvWin_Click
   
   
End Sub

Public Sub ShowAbout()
    'frmAbout.Show vbModal
End Sub



Public Sub CloseAll()
    Dim Form As Form
   For Each Form In Forms
      ''Unload all active childs
      If Form.Name <> Me.Name And Form.Name <> "frmShortcuts" Then Unload Form
   Next Form
End Sub


