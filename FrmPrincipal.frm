VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.MDIForm FrmPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "Projeto 3º Bimestre"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6945
   Icon            =   "FrmPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "FrmPrincipal.frx":030A
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   13
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Clientes"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Livros"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Vendas"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Relatórios"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
      MouseIcon       =   "FrmPrincipal.frx":058C
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   6330
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327680
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   7
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   2778
            Text            =   "Anselmo / Fábio "
            TextSave        =   "Anselmo / Fábio "
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Componentes Equipe"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "CAPS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "NUM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "SCRL"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "INS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   2778
            TextSave        =   "28/10/1999"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "20:24"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      MouseIcon       =   "FrmPrincipal.frx":05A8
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5880
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmPrincipal.frx":05C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmPrincipal.frx":08DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmPrincipal.frx":0BF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmPrincipal.frx":0F12
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmPrincipal.frx":122C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmPrincipal.frx":1546
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmPrincipal.frx":1860
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmPrincipal.frx":1B7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmPrincipal.frx":1E94
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmPrincipal.frx":21AE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuControles 
      Caption         =   "&Controles"
      Begin VB.Menu MnuClientes 
         Caption         =   "&Clientes"
      End
      Begin VB.Menu MnuLivros 
         Caption         =   "&Livros"
      End
      Begin VB.Menu MnuVendas 
         Caption         =   "&Vendas"
      End
      Begin VB.Menu mnuSpe1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRelatorios 
         Caption         =   "&Relatórios"
      End
   End
   Begin VB.Menu MnuObjetos 
      Caption         =   "&Objetos"
   End
   Begin VB.Menu MnuJan 
      Caption         =   "&Janelas"
      WindowList      =   -1  'True
      Begin VB.Menu MnuCas 
         Caption         =   "&Cascata"
      End
      Begin VB.Menu MnuHor 
         Caption         =   "&Horizontal"
      End
      Begin VB.Menu MnuVer 
         Caption         =   "&Vertical"
      End
      Begin VB.Menu MnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMinAll 
         Caption         =   "Mi&nimizar Todas"
      End
      Begin VB.Menu MnuRest 
         Caption         =   "&Restaurar Todas"
      End
      Begin VB.Menu MnuFechar 
         Caption         =   "&Fechar Todas"
      End
   End
   Begin VB.Menu MnuSair 
      Caption         =   "&Sair"
      Begin VB.Menu MnuSobre 
         Caption         =   "&Sobre"
      End
      Begin VB.Menu MnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFim 
         Caption         =   "&Finalizar"
      End
   End
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub MnuCas_Click()
   FrmMain.Arrange vbCascade
End Sub



Private Sub MnuClientes_Click()
   frmClientes.Show
End Sub

Private Sub MnuFechar_Click()
   Dim i As Form
   For Each i In Forms
      If i.Name <> "FrmMain" Then Unload i
   Next i
   MnuAlg.Checked = False
   MnuEst.Checked = False
   MnuVar.Checked = False
   MnuOpe.Checked = False
End Sub

Private Sub MnuFim_Click()
   Unload Me
End Sub

Private Sub MnuHor_Click()
   FrmMain.Arrange vbTileHorizontal
End Sub

Private Sub MnuMinAll_Click()
   Dim i As Form
   For Each i In Forms
      If i.Name <> "FrmMain" Then i.WindowState = 1
   Next i
End Sub

Private Sub MnuOpe_Click()
   If MnuOpe.Checked = False Then
      MnuOpe.Checked = True
      FrmOperadores.Show
   Else
      MnuOpe.Checked = False
      Unload FrmOperadores
   End If
End Sub

Private Sub MnuRest_Click()
   Dim i As Form
   For Each i In Forms
      If i.Name <> "FrmMain" Then i.WindowState = 0
   Next i
End Sub

Private Sub MnuSobre_Click()
   FrmSobre.Show
End Sub

Private Sub MnuVar_Click()
   If MnuVar.Checked = False Then
      MnuVar.Checked = True
      FrmMemoria.Show
   Else
      MnuVar.Checked = False
      Unload FrmMemoria
   End If
End Sub

Private Sub MnuVer_Click()
   FrmMain.Arrange vbTileVertical
End Sub

Private Sub StbMain_PanelClick(ByVal Panel As ComctlLib.Panel)
   If Panel.Index = 1 Then FrmSobre.Show
End Sub

Private Sub TobMain_ButtonClick(ByVal Button As ComctlLib.Button)
   Select Case Button.Index
      Case 1
         If MnuAlg.Checked = False Then
            FrmAlg.Show
            MnuAlg.Checked = True
         Else
            Unload FrmAlg
            MnuAlg.Checked = False
         End If
      Case 3
         If MnuVar.Checked = False Then
            FrmMemoria.Show
            MnuVar.Checked = True
         Else
            Unload FrmMemoria
            MnuVar.Checked = False
         End If
      Case 4
         If MnuEst.Checked = False Then
            FrmEstCont.Show
            MnuEst.Checked = True
         Else
            Unload FrmEstCont
            MnuEst.Checked = False
         End If
      Case 5
         If MnuOpe.Checked = False Then
            FrmOperadores.Show
            MnuOpe.Checked = True
         Else
            Unload FrmOperadores
            MnuOpe.Checked = False
         End If
      Case 7
         Dim i As Form
         For Each i In Forms
            If i.Name <> "FrmMain" Then Unload i
         Next i
         MnuAlg.Checked = False
         MnuEst.Checked = False
         MnuVar.Checked = False
         MnuOpe.Checked = False
      Case 9
         Unload Me
   End Select
End Sub
