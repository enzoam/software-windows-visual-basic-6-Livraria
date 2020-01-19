VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmVendas 
   Caption         =   "Cadastro de Vendas"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   6000
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   6000
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   960
         Width           =   1335
      End
      Begin MSDBCtls.DBCombo DBCombo2 
         Height          =   315
         Left            =   2160
         TabIndex        =   18
         Top             =   1440
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   327680
         Text            =   "DBCombo2"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   315
         Left            =   2160
         TabIndex        =   17
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   327680
         Text            =   "DBCombo1"
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2160
         TabIndex        =   8
         Top             =   1920
         Width           =   3615
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   2640
         Width           =   3615
      End
      Begin VB.TextBox Text6 
         Height          =   405
         Left            =   2160
         TabIndex        =   6
         Top             =   3120
         Width           =   3615
      End
      Begin VB.CommandButton btnadd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   3720
         Width           =   1200
      End
      Begin VB.CommandButton btndelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   3720
         Width           =   1200
      End
      Begin VB.CommandButton btnrefresh 
         Caption         =   "&Reflesh"
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   3720
         Width           =   1200
      End
      Begin VB.CommandButton btnupdate 
         Caption         =   "&Update"
         Height          =   375
         Left            =   4440
         TabIndex        =   2
         Top             =   3720
         Width           =   1200
      End
      Begin VB.CommandButton btnclose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   5760
         TabIndex        =   1
         Top             =   3720
         Width           =   1200
      End
      Begin VB.Label Label2 
         Caption         =   "Preço :"
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Data Venda :"
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Quantidade :"
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Forma Pagto :"
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Data Pagto :"
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Livro :"
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label8 
         Caption         =   "Cliente :"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label2_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub Text6_Change()

End Sub
