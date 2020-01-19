VERSION 5.00
Begin VB.Form frmlivros 
   Caption         =   "Cadastro de Livros"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   7335
      Begin VB.TextBox txtco 
         DataField       =   "Codigo"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtno 
         DataField       =   "Nome"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   720
         Width           =   4455
      End
      Begin VB.TextBox txtau 
         DataField       =   "Autor"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1560
         TabIndex        =   9
         Top             =   1200
         Width           =   4455
      End
      Begin VB.TextBox txted 
         DataField       =   "Editora"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   1680
         Width           =   4455
      End
      Begin VB.TextBox txtpr 
         DataField       =   "Preco"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox txtqu 
         DataField       =   "Quantidade"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CommandButton btnno 
         Caption         =   "&Novo"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton btnre 
         Caption         =   "&Remover"
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton btnat 
         Caption         =   "&Atualizar"
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton btnsa 
         Caption         =   "&Salvar"
         Height          =   375
         Left            =   4560
         TabIndex        =   2
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton btnfe 
         Caption         =   "&Fechar"
         Height          =   375
         Left            =   6000
         TabIndex        =   1
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Data Data1 
         Connect         =   "Access"
         DatabaseName    =   "K:\Users\Info\204485\Livraria.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Livros"
         Top             =   3960
         Width           =   6855
      End
      Begin VB.Label lblco 
         Caption         =   "Código :"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label lblno 
         Caption         =   "Nome :"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label lblau 
         Caption         =   "Autor :"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label lbled 
         Caption         =   "Editora :"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   3615
      End
      Begin VB.Label lblpr 
         Caption         =   "Preco :"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   2280
         Width           =   4215
      End
      Begin VB.Label lblqu 
         Caption         =   "Quantidade :"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   2760
         Width           =   4455
      End
   End
End
Attribute VB_Name = "frmlivros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnfe_Click()
      frmlivros.Hide
End Sub

Private Sub btnno_Click()
      Data1.Recordset.AddNew
End Sub

Private Sub btnsa_Click()
    Data1.Recordset.Update
End Sub
