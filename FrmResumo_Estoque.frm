VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmResumo_Estoque 
   Caption         =   "Resumo do Estoque"
   ClientHeight    =   3000
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
   Icon            =   "FrmResumo_Estoque.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
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
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   6735
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Estoque.frx":0CCA
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Estoque.frx":0D32
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTipoProd 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Estoque.frx":0DA4
         TabIndex        =   5
         Top             =   600
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Estoque.frx":0E08
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblUltPed 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Estoque.frx":0E78
         TabIndex        =   7
         Top             =   960
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblQtdeEst 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Estoque.frx":0EE0
         TabIndex        =   8
         Top             =   1680
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Estoque.frx":0F4C
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Estoque.frx":0FBC
         TabIndex        =   10
         Top             =   1680
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblQtdeMin 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Estoque.frx":102E
         TabIndex        =   11
         Top             =   1320
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblProduto 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Estoque.frx":1096
         TabIndex        =   12
         Top             =   240
         Width           =   5175
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
      Top             =   2160
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
         OleObjectBlob   =   "FrmResumo_Estoque.frx":10F6
         Top             =   120
      End
   End
End
Attribute VB_Name = "FrmResumo_Estoque"
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
  FrmResumo_Estoque.Left = (MDIPrincipal.Width / 2) - (FrmResumo_Estoque.Width / 2)
  FrmResumo_Estoque.Top = (MDIPrincipal.Height / 3) - (FrmResumo_Estoque.Height / 3)
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
        
    Height = 3510
    Width = 7095
    
    Call MontaResumo
    
    MDIPrincipal.Enabled = False
End Sub

Sub MontaResumo()
        
    FrmPrincipal.GridEstoque.Row = FrmPrincipal.GridEstoque.ActiveRow
    
    'Produto
    FrmPrincipal.GridEstoque.Col = 1
    LblProduto.Caption = FrmPrincipal.GridEstoque.Text

    'Tipo Produto
    FrmPrincipal.GridEstoque.Col = 2
    LblTipoProd.Caption = FrmPrincipal.GridEstoque.Text

    'Último pedido
    FrmPrincipal.GridEstoque.Col = 3
    LblUltPed.Caption = FrmPrincipal.GridEstoque.Text

    'Qtde Mínima
    FrmPrincipal.GridEstoque.Col = 4
    LblQtdeMin.Caption = FrmPrincipal.GridEstoque.Text

    'Qtde em estoque
    FrmPrincipal.GridEstoque.Col = 5
    LblQtdeEst.Caption = FrmPrincipal.GridEstoque.Text

End Sub
