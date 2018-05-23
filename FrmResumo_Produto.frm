VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmResumo_Produto 
   Caption         =   "Resumo do Produto"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmResumo_Produto.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   6975
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   6735
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Produto.frx":0CCA
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Produto.frx":0D32
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblFornecedor 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Produto.frx":0DA0
         TabIndex        =   5
         Top             =   600
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Produto.frx":0E04
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTipoProd 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Produto.frx":0E76
         TabIndex        =   7
         Top             =   960
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblPrecoForn 
         Height          =   255
         Left            =   1800
         OleObjectBlob   =   "FrmResumo_Produto.frx":0EDE
         TabIndex        =   8
         Top             =   1680
         Width           =   4815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Produto.frx":0F4A
         TabIndex        =   9
         Top             =   1320
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Produto.frx":0FC6
         TabIndex        =   10
         Top             =   1680
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblDescProd 
         Height          =   255
         Left            =   1920
         OleObjectBlob   =   "FrmResumo_Produto.frx":1040
         TabIndex        =   11
         Top             =   1320
         Width           =   4695
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblProduto 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Produto.frx":10A8
         TabIndex        =   12
         Top             =   240
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblPrecoVendaUnit 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "FrmResumo_Produto.frx":1108
         TabIndex        =   13
         Top             =   2040
         Width           =   4575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Produto.frx":1174
         TabIndex        =   14
         Top             =   2040
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblPrecoVendaAtac 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "FrmResumo_Produto.frx":11F6
         TabIndex        =   15
         Top             =   2400
         Width           =   4575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Produto.frx":1262
         TabIndex        =   16
         Top             =   2400
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblMoeda 
         Height          =   255
         Left            =   840
         OleObjectBlob   =   "FrmResumo_Produto.frx":12E2
         TabIndex        =   17
         Top             =   2760
         Width           =   5775
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Produto.frx":134E
         TabIndex        =   18
         Top             =   2760
         Width           =   1215
      End
   End
   Begin VB.Frame FraBotaoCli 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   6735
      Begin VB.CommandButton CmdFechar 
         Caption         =   "&Fechar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   0
         ToolTipText     =   "Fechar"
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   360
         OleObjectBlob   =   "FrmResumo_Produto.frx":13B2
         Top             =   120
      End
   End
End
Attribute VB_Name = "FrmResumo_Produto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdFechar_Click()
    Unload Me
    
    MDIPrincipal.Enabled = True
    MDIPrincipal.WindowState = 2
    MDIPrincipal.SetFocus
End Sub

Private Sub Form_Resize()
  FrmResumo_Produto.Left = (MDIPrincipal.Width / 2) - (FrmResumo_Produto.Width / 2)
  FrmResumo_Produto.Top = (MDIPrincipal.Height / 3) - (FrmResumo_Produto.Height / 3)
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
        
    Height = 4575
    Width = 7095
    
    Call MontaResumo
    
    MDIPrincipal.Enabled = False
End Sub

Sub MontaResumo()
        
    FrmPrincipal.GridProduto.Row = FrmPrincipal.GridProduto.ActiveRow
    
    'Produto
    FrmPrincipal.GridProduto.Col = 1
    LblProduto.Caption = FrmPrincipal.GridProduto.Text

    'Fornecedor
    FrmPrincipal.GridProduto.Col = 2
    LblFornecedor.Caption = FrmPrincipal.GridProduto.Text

    'Tipo produto
    FrmPrincipal.GridProduto.Col = 3
    LblTipoProd.Caption = FrmPrincipal.GridProduto.Text

    'Descrição de produto
    FrmPrincipal.GridProduto.Col = 4
    LblDescProd.Caption = FrmPrincipal.GridProduto.Text

    'Preço do Fornecedor
    FrmPrincipal.GridProduto.Col = 5
    LblPrecoForn.Caption = FrmPrincipal.GridProduto.Text

    'Preço de venda unitário
    FrmPrincipal.GridProduto.Col = 6
    LblPrecoVendaUnit.Caption = FrmPrincipal.GridProduto.Text

    'Preço de venda atacado
    FrmPrincipal.GridProduto.Col = 7
    LblPrecoVendaAtac.Caption = FrmPrincipal.GridProduto.Text

    'Moeda
    FrmPrincipal.GridProduto.Col = 8
    LblMoeda.Caption = FrmPrincipal.GridProduto.Text

End Sub
