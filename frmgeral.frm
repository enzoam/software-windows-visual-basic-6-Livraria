VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmgeral 
   Caption         =   "CATÁLOGOS DE CLIENTES - Dados Gerais"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.Data datcadastro 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Trabalho\TABELAS\Catalogo.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4200
      Width           =   1335
   End
   Begin MSDBGrid.DBGrid dbggeral 
      Bindings        =   "frmgeral.frx":0000
      Height          =   3975
      Left            =   0
      OleObjectBlob   =   "frmgeral.frx":0016
      TabIndex        =   0
      Top             =   0
      Width           =   10575
   End
End
Attribute VB_Name = "frmgeral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
