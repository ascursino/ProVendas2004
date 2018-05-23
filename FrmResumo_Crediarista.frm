VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmResumo_Crediarista 
   Caption         =   "Resumo do Crediarista"
   ClientHeight    =   5985
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
   Icon            =   "FrmResumo_Crediarista.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
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
      Height          =   5055
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   6735
      Begin VB.TextBox TxtObs 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3840
         Width           =   5295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Crediarista.frx":0CCA
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Crediarista.frx":0D2C
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Crediarista.frx":0D96
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Crediarista.frx":0DFC
         TabIndex        =   6
         Top             =   1680
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Crediarista.frx":0E62
         TabIndex        =   7
         Top             =   2400
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Crediarista.frx":0EC6
         TabIndex        =   8
         Top             =   3120
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Crediarista.frx":0F26
         TabIndex        =   9
         Top             =   1320
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Crediarista.frx":0F86
         TabIndex        =   10
         Top             =   2040
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Crediarista.frx":0FEE
         TabIndex        =   11
         Top             =   2760
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Crediarista.frx":1058
         TabIndex        =   12
         Top             =   3480
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Crediarista.frx":10BC
         TabIndex        =   13
         Top             =   3840
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblBairro 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Crediarista.frx":112A
         TabIndex        =   14
         Top             =   960
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCpf 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Crediarista.frx":1192
         TabIndex        =   15
         Top             =   3120
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCidade 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Crediarista.frx":11FA
         TabIndex        =   16
         Top             =   1680
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblEndereco 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Crediarista.frx":1262
         TabIndex        =   17
         Top             =   600
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNome 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Crediarista.frx":12CA
         TabIndex        =   18
         Top             =   240
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNasc 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Crediarista.frx":1332
         TabIndex        =   19
         Top             =   2400
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTel 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Crediarista.frx":139A
         TabIndex        =   20
         Top             =   2760
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCep 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Crediarista.frx":1402
         TabIndex        =   21
         Top             =   1320
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblEstado 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Crediarista.frx":146A
         TabIndex        =   22
         Top             =   2040
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblEmail 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Crediarista.frx":14D2
         TabIndex        =   23
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
      Top             =   5160
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
         OleObjectBlob   =   "FrmResumo_Crediarista.frx":153A
         Top             =   120
      End
   End
End
Attribute VB_Name = "FrmResumo_Crediarista"
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
  FrmResumo_Crediarista.Left = (MDIPrincipal.Width / 2) - (FrmResumo_Crediarista.Width / 2)
  FrmResumo_Crediarista.Top = (MDIPrincipal.Height / 3) - (FrmResumo_Crediarista.Height / 3)
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
        
    Height = 6495
    Width = 7065
    
    Call MontaResumo
    
    MDIPrincipal.Enabled = False
End Sub

Sub MontaResumo()
        
    FrmPrincipal.GridCredsta.Row = FrmPrincipal.GridCredsta.ActiveRow
    
    'Nome
    FrmPrincipal.GridCredsta.Col = 1
    LblNome.Caption = FrmPrincipal.GridCredsta.Text

    'Endereço
    FrmPrincipal.GridCredsta.Col = 2
    LblEndereco.Caption = FrmPrincipal.GridCredsta.Text

    'Bairro
    FrmPrincipal.GridCredsta.Col = 3
    LblBairro.Caption = FrmPrincipal.GridCredsta.Text

    'Cep
    FrmPrincipal.GridCredsta.Col = 4
    LblCep.Caption = FrmPrincipal.GridCredsta.Text

    'Cidade
    FrmPrincipal.GridCredsta.Col = 5
    LblCidade.Caption = FrmPrincipal.GridCredsta.Text

    'Estado
    FrmPrincipal.GridCredsta.Col = 6
    LblEstado.Caption = FrmPrincipal.GridCredsta.Text

    'Data Nascimento
    FrmPrincipal.GridCredsta.Col = 7
    LblNasc.Caption = FrmPrincipal.GridCredsta.Text

    'Telefone
    FrmPrincipal.GridCredsta.Col = 8
    LblTel.Caption = FrmPrincipal.GridCredsta.Text

    'Cpf
    FrmPrincipal.GridCredsta.Col = 9
    LblCpf.Caption = FrmPrincipal.GridCredsta.Text

    'Email
    FrmPrincipal.GridCredsta.Col = 10
    LblEmail.Caption = FrmPrincipal.GridCredsta.Text

    'Observação
    FrmPrincipal.GridCredsta.Col = 11
    TxtObs.Text = FrmPrincipal.GridCredsta.Text
    
End Sub

