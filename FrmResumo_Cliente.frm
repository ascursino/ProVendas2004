VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmResumo_Cliente 
   Caption         =   "Resumo do Cliente"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6945
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
   Icon            =   "FrmResumo_Cliente.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   6945
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
      Height          =   6375
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   6735
      Begin ACTIVESKINLibCtl.SkinLabel LblObs 
         Height          =   855
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Cliente.frx":0CCA
         TabIndex        =   32
         Top             =   5280
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Cliente.frx":0D32
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Cliente.frx":0D94
         TabIndex        =   4
         Top             =   1320
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Cliente.frx":0DFE
         TabIndex        =   5
         Top             =   1680
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Cliente.frx":0E64
         TabIndex        =   6
         Top             =   2400
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Cliente.frx":0ECA
         TabIndex        =   7
         Top             =   3120
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Cliente.frx":0F2E
         TabIndex        =   8
         Top             =   4560
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Cliente.frx":0F8E
         TabIndex        =   9
         Top             =   2040
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Cliente.frx":0FEE
         TabIndex        =   10
         Top             =   2760
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Cliente.frx":1056
         TabIndex        =   11
         Top             =   3480
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Cliente.frx":10C0
         TabIndex        =   12
         Top             =   3840
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Cliente.frx":1128
         TabIndex        =   13
         Top             =   4920
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Cliente.frx":118C
         TabIndex        =   14
         Top             =   5280
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Cliente.frx":11FA
         TabIndex        =   15
         Top             =   960
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Cliente.frx":125C
         TabIndex        =   16
         Top             =   4200
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblBairro 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Cliente.frx":12BC
         TabIndex        =   17
         Top             =   1680
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblSexo 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Cliente.frx":1324
         TabIndex        =   18
         Top             =   960
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCpf 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Cliente.frx":138C
         TabIndex        =   19
         Top             =   4560
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCidade 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Cliente.frx":13F4
         TabIndex        =   20
         Top             =   2400
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblEndereco 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Cliente.frx":145C
         TabIndex        =   21
         Top             =   1320
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNome 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Cliente.frx":14C4
         TabIndex        =   22
         Top             =   240
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNasc 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Cliente.frx":152C
         TabIndex        =   23
         Top             =   3120
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCel 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Cliente.frx":1594
         TabIndex        =   24
         Top             =   3840
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTel 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Cliente.frx":15FC
         TabIndex        =   25
         Top             =   3480
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCep 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Cliente.frx":1664
         TabIndex        =   26
         Top             =   2040
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblEstado 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Cliente.frx":16CC
         TabIndex        =   27
         Top             =   2760
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblFax 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Cliente.frx":1734
         TabIndex        =   28
         Top             =   4200
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblEmail 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Cliente.frx":179C
         TabIndex        =   29
         Top             =   4920
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Cliente.frx":1804
         TabIndex        =   30
         Top             =   600
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblClidesde 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Cliente.frx":1878
         TabIndex        =   31
         Top             =   600
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
      Top             =   6480
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
         OleObjectBlob   =   "FrmResumo_Cliente.frx":18E0
         Top             =   120
      End
   End
End
Attribute VB_Name = "FrmResumo_Cliente"
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
  FrmResumo_Cliente.Left = (MDIPrincipal.Width / 2) - (FrmResumo_Cliente.Width / 2)
  FrmResumo_Cliente.Top = (MDIPrincipal.Height / 3) - (FrmResumo_Cliente.Height / 3)
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
        
    Height = 7815
    Width = 7065
    
    Call MontaResumo
    
    MDIPrincipal.Enabled = False
End Sub

Sub MontaResumo()
        
    FrmPrincipal.GridCliente.Row = FrmPrincipal.GridCliente.ActiveRow
    
    'Nome
    FrmPrincipal.GridCliente.Col = 1
    LblNome.Caption = FrmPrincipal.GridCliente.Text

    'Cliente desde
    FrmPrincipal.GridCliente.Col = 2
    LblClidesde.Caption = FrmPrincipal.GridCliente.Text

    'Sexo
    FrmPrincipal.GridCliente.Col = 3
    LblSexo.Caption = FrmPrincipal.GridCliente.Text

    'Endereço
    FrmPrincipal.GridCliente.Col = 4
    LblEndereco.Caption = FrmPrincipal.GridCliente.Text

    'Bairro
    FrmPrincipal.GridCliente.Col = 5
    LblBairro.Caption = FrmPrincipal.GridCliente.Text

    'Cep
    FrmPrincipal.GridCliente.Col = 6
    LblCep.Caption = FrmPrincipal.GridCliente.Text

    'Cidade
    FrmPrincipal.GridCliente.Col = 7
    LblCidade.Caption = FrmPrincipal.GridCliente.Text

    'Estado
    FrmPrincipal.GridCliente.Col = 8
    LblEstado.Caption = FrmPrincipal.GridCliente.Text

    'Data Nascimento
    FrmPrincipal.GridCliente.Col = 9
    LblNasc.Caption = FrmPrincipal.GridCliente.Text

    'Telefone
    FrmPrincipal.GridCliente.Col = 10
    LblTel.Caption = FrmPrincipal.GridCliente.Text

    'Celular
    FrmPrincipal.GridCliente.Col = 11
    LblCel.Caption = FrmPrincipal.GridCliente.Text

    'Fax
    FrmPrincipal.GridCliente.Col = 12
    LblFax.Caption = FrmPrincipal.GridCliente.Text

    'Cpf
    FrmPrincipal.GridCliente.Col = 13
    LblCpf.Caption = FrmPrincipal.GridCliente.Text

    'Email
    FrmPrincipal.GridCliente.Col = 14
    LblEmail.Caption = FrmPrincipal.GridCliente.Text

    'Observação
    FrmPrincipal.GridCliente.Col = 15
    LblObs.Caption = FrmPrincipal.GridCliente.Text
    
End Sub

