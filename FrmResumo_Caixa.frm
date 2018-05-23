VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmResumo_Caixa 
   Caption         =   "Resumo do Movimento de Caixa"
   ClientHeight    =   3705
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
   Icon            =   "FrmResumo_Caixa.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
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
      Height          =   2775
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   6735
      Begin VB.TextBox TxtDescrMov 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Text            =   "?????"
         Top             =   1200
         Width           =   6495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Caixa.frx":0CCA
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblDtMov 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "FrmResumo_Caixa.frx":0D36
         TabIndex        =   4
         Top             =   2040
         Width           =   4575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Caixa.frx":0DA2
         TabIndex        =   5
         Top             =   2040
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTipoMov 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "FrmResumo_Caixa.frx":0E1E
         TabIndex        =   6
         Top             =   600
         Width           =   4575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Caixa.frx":0E8A
         TabIndex        =   7
         Top             =   600
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblValorMov 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "FrmResumo_Caixa.frx":0F06
         TabIndex        =   8
         Top             =   2400
         Width           =   4575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Caixa.frx":0F72
         TabIndex        =   9
         Top             =   2400
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCodVenda 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Caixa.frx":0FD6
         TabIndex        =   10
         Top             =   240
         Width           =   4575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Caixa.frx":1042
         TabIndex        =   11
         Top             =   240
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
      Top             =   2880
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
         OleObjectBlob   =   "FrmResumo_Caixa.frx":10B0
         Top             =   120
      End
   End
End
Attribute VB_Name = "FrmResumo_Caixa"
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
  FrmResumo_Caixa.Left = (MDIPrincipal.Width / 2) - (FrmResumo_Caixa.Width / 2)
  FrmResumo_Caixa.Top = (MDIPrincipal.Height / 3) - (FrmResumo_Caixa.Height / 3)
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
        
    Height = 4515
    Width = 7095
    
    Call MontaResumo
    
    MDIPrincipal.Enabled = False
End Sub

Sub MontaResumo()
        
    FrmPrincipal.GridCaixa.Row = FrmPrincipal.GridCaixa.ActiveRow
    
    'Cód. Venda
    FrmPrincipal.GridCaixa.Col = 1
    LblCodVenda.Caption = FrmPrincipal.GridCaixa.Text
    
    'Descrição
    FrmPrincipal.GridCaixa.Col = 2
    TxtDescrMov.Text = FrmPrincipal.GridCaixa.Text
    
    'Data do movimento
    FrmPrincipal.GridCaixa.Col = 3
    LblDtMov.Caption = FrmPrincipal.GridCaixa.Text
    
    'Tipo de Movimento
    FrmPrincipal.GridCaixa.Col = 4
    LblTipoMov.Caption = FrmPrincipal.GridCaixa.Text

    'Valor
    FrmPrincipal.GridCaixa.Col = 5
    If FrmPrincipal.GridCaixa.Text <> "" Then
        LblValorMov.Caption = FrmPrincipal.GridCaixa.Text
    Else
        FrmPrincipal.GridCaixa.Col = 6
        LblValorMov.Caption = FrmPrincipal.GridCaixa.Text
    End If
        
End Sub
