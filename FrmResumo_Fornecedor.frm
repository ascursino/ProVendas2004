VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmResumo_Fornecedor 
   Caption         =   "Resumo do Fornecedor"
   ClientHeight    =   6960
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
   Icon            =   "FrmResumo_Fornecedor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
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
      Height          =   6015
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   6735
      Begin ACTIVESKINLibCtl.SkinLabel LblObs 
         Height          =   855
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":0CCA
         TabIndex        =   3
         Top             =   4920
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":0D32
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":0DA0
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":0E0A
         TabIndex        =   6
         Top             =   1320
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":0E70
         TabIndex        =   7
         Top             =   2040
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":0ED6
         TabIndex        =   8
         Top             =   2760
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":0F38
         TabIndex        =   9
         Top             =   1680
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":0F98
         TabIndex        =   10
         Top             =   2400
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":1000
         TabIndex        =   11
         Top             =   3840
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":106A
         TabIndex        =   12
         Top             =   4200
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":10D2
         TabIndex        =   13
         Top             =   3120
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":1136
         TabIndex        =   14
         Top             =   4920
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":11A4
         TabIndex        =   15
         Top             =   4560
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblBairro 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":1204
         TabIndex        =   16
         Top             =   1320
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCnpj 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":126C
         TabIndex        =   17
         Top             =   2760
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCidade 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":12D4
         TabIndex        =   18
         Top             =   2040
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblEndereco 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":133C
         TabIndex        =   19
         Top             =   960
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblForn 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":13A4
         TabIndex        =   20
         Top             =   240
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCel 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":140C
         TabIndex        =   21
         Top             =   4200
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTel 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":1474
         TabIndex        =   22
         Top             =   3840
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCep 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":14DC
         TabIndex        =   23
         Top             =   1680
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblEstado 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":1544
         TabIndex        =   24
         Top             =   2400
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblFax 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":15AC
         TabIndex        =   25
         Top             =   4560
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblEmail 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":1614
         TabIndex        =   26
         Top             =   3120
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":167C
         TabIndex        =   27
         Top             =   600
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTipo 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":16DE
         TabIndex        =   28
         Top             =   600
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":1746
         TabIndex        =   29
         Top             =   3480
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblResp 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":17B6
         TabIndex        =   30
         Top             =   3480
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
      Top             =   6120
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
         OleObjectBlob   =   "FrmResumo_Fornecedor.frx":181E
         Top             =   120
      End
   End
End
Attribute VB_Name = "FrmResumo_Fornecedor"
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
  FrmResumo_Fornecedor.Left = (MDIPrincipal.Width / 2) - (FrmResumo_Fornecedor.Width / 2)
  FrmResumo_Fornecedor.Top = (MDIPrincipal.Height / 3) - (FrmResumo_Fornecedor.Height / 3)
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
        
    Height = 7470
    Width = 7095
    
    Call MontaResumo
    
    MDIPrincipal.Enabled = False
End Sub

Sub MontaResumo()
        
    FrmPrincipal.GridFornecedor.Row = FrmPrincipal.GridFornecedor.ActiveRow
    
    'Fornecedor
    FrmPrincipal.GridFornecedor.Col = 1
    LblForn.Caption = FrmPrincipal.GridFornecedor.Text

    'Tipo
    FrmPrincipal.GridFornecedor.Col = 2
    LblTipo.Caption = FrmPrincipal.GridFornecedor.Text

    'Endereço
    FrmPrincipal.GridFornecedor.Col = 3
    LblEndereco.Caption = FrmPrincipal.GridFornecedor.Text

    'Bairro
    FrmPrincipal.GridFornecedor.Col = 4
    LblBairro.Caption = FrmPrincipal.GridFornecedor.Text

    'Cep
    FrmPrincipal.GridFornecedor.Col = 5
    LblCep.Caption = FrmPrincipal.GridFornecedor.Text

    'Cidade
    FrmPrincipal.GridFornecedor.Col = 6
    LblCidade.Caption = FrmPrincipal.GridFornecedor.Text

    'Estado
    FrmPrincipal.GridFornecedor.Col = 7
    LblEstado.Caption = FrmPrincipal.GridFornecedor.Text

    'CNPJ
    FrmPrincipal.GridFornecedor.Col = 8
    LblCnpj.Caption = FrmPrincipal.GridFornecedor.Text

    'Email
    FrmPrincipal.GridFornecedor.Col = 9
    LblEmail.Caption = FrmPrincipal.GridFornecedor.Text

    'Responsável
    FrmPrincipal.GridFornecedor.Col = 10
    LblResp.Caption = FrmPrincipal.GridFornecedor.Text

    'Telefone
    FrmPrincipal.GridFornecedor.Col = 11
    LblTel.Caption = FrmPrincipal.GridFornecedor.Text

    'Celular
    FrmPrincipal.GridFornecedor.Col = 12
    LblCel.Caption = FrmPrincipal.GridFornecedor.Text

    'Fax
    FrmPrincipal.GridFornecedor.Col = 13
    LblFax.Caption = FrmPrincipal.GridFornecedor.Text

    'Observação
    FrmPrincipal.GridFornecedor.Col = 14
    LblObs.Caption = FrmPrincipal.GridFornecedor.Text

End Sub

