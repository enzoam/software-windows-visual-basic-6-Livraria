VERSION 5.00
Begin VB.Form frmClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   8190
   Begin VB.Frame Frame1 
      Height          =   5895
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   7815
      Begin VB.CommandButton btnirpara 
         Caption         =   "Ir Para"
         Height          =   495
         Left            =   5760
         TabIndex        =   10
         Top             =   4560
         Width           =   1575
      End
      Begin VB.CommandButton btnexcluir 
         Caption         =   "Excluir"
         Height          =   495
         Left            =   5760
         TabIndex        =   9
         Top             =   3720
         Width           =   1575
      End
      Begin VB.CommandButton btnsalvar 
         Caption         =   "Salvar"
         Height          =   495
         Left            =   5760
         TabIndex        =   8
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton btnnovo 
         Caption         =   "Novo"
         Height          =   495
         Left            =   5760
         TabIndex        =   7
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton btnfechar 
         Caption         =   "Fechar"
         Height          =   495
         Left            =   5760
         TabIndex        =   6
         Top             =   480
         Width           =   1575
      End
      Begin VB.Frame Frmsub 
         Caption         =   "Ordenado Por"
         Height          =   1335
         Left            =   600
         TabIndex        =   16
         Top             =   3720
         Width           =   4695
         Begin VB.OptionButton optnome 
            Caption         =   "Nome"
            Height          =   375
            Left            =   2880
            TabIndex        =   5
            Top             =   480
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton optcodigo 
            Caption         =   "Codigo"
            Height          =   255
            Left            =   840
            TabIndex        =   4
            Top             =   600
            Width           =   1455
         End
      End
      Begin VB.TextBox Txtcodigo 
         DataField       =   "Codigo"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   2520
         TabIndex        =   0
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox txtnome 
         DataField       =   "Nome"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   2520
         TabIndex        =   1
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtendereco 
         DataField       =   "Endereco"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txttelefone 
         DataField       =   "Telefone"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   2880
         Width           =   2775
      End
      Begin VB.Data Data1 
         Connect         =   "Access"
         DatabaseName    =   "K:\Users\Info\204485\Livraria.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Clientes"
         Top             =   5160
         Width           =   4695
      End
      Begin VB.Label lblcodigo 
         Caption         =   "Código : "
         Height          =   495
         Left            =   600
         TabIndex        =   15
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblnome 
         Caption         =   "Nome :"
         Height          =   495
         Left            =   600
         TabIndex        =   14
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label lblendereco 
         Caption         =   "Endereço :"
         Height          =   495
         Left            =   600
         TabIndex        =   13
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label lbltelefone 
         Caption         =   "Telefone :"
         Height          =   495
         Left            =   600
         TabIndex        =   12
         Top             =   3000
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Text1_Change()

End Sub

Private Sub btnexcluir_Click()
   Data1.Recordset .Delete
   If Not Data1.Recordset.EOF Then
      Data1.Recordset.MoveFirst
   End If
End Sub

Private Sub btnfechar_Click()
   frmClientes.Hide
End Sub

Private Sub btnirpara_Click()
   Dim nome_cli, expr_nome, resposta As String
   nome_cli = InputBox("Nome do Cliente:", "Ir Para")
   expr_nome = "Nome='" & nome_cli & "'"
   Data1.Recordset.FindFirst expr_nome
   If Data1.Recordset.NoMatch Then
      resposta = MsgBox("Cliente não encontrado!")
   End If
   
End Sub

Private Sub btnnovo_Click()
   Data1.Recordset.AddNew
End Sub

Private Sub btnsalvar_Click()
   Data1.Recordset.Update
End Sub

Private Sub optcodigo_Click()
   If optcodigo.Value Then
      Data1.RecordSource = "Select * from Clientes order by codigo"
      Data1.Refresh
   End If
End Sub

Private Sub optnome_Click()
   If optnome.Value Then
      Data1.RecordSource = "Select * from Clientes order by Nome"
      Data1.Refresh
   End If
End Sub
