VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmResumo_Vendedor 
   Caption         =   "Resumo do Vendedor"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6990
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
   Icon            =   "FrmResumo_Vendedor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   6990
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
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   6735
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Vendedor.frx":0CCA
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNome 
         Height          =   255
         Left            =   1080
         OleObjectBlob   =   "FrmResumo_Vendedor.frx":0D2C
         TabIndex        =   4
         Top             =   240
         Width           =   5535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Vendedor.frx":0D94
         TabIndex        =   5
         Top             =   600
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTel 
         Height          =   255
         Left            =   1080
         OleObjectBlob   =   "FrmResumo_Vendedor.frx":0DFE
         TabIndex        =   6
         Top             =   600
         Width           =   5535
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
      Top             =   1080
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
         OleObjectBlob   =   "FrmResumo_Vendedor.frx":0E66
         Top             =   120
      End
   End
End
Attribute VB_Name = "FrmResumo_Vendedor"
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
  FrmResumo_Vendedor.Left = (MDIPrincipal.Width / 2) - (FrmResumo_Vendedor.Width / 2)
  FrmResumo_Vendedor.Top = (MDIPrincipal.Height / 3) - (FrmResumo_Vendedor.Height / 3)
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
        
    Height = 2430
    Width = 7110
    
    Call MontaResumo
    
    MDIPrincipal.Enabled = False
End Sub

Sub MontaResumo()
        
    FrmPrincipal.GridVendedor.Row = FrmPrincipal.GridVendedor.ActiveRow
    
    'Nome
    FrmPrincipal.GridVendedor.Col = 1
    LblNome.Caption = FrmPrincipal.GridVendedor.Text

    'Telefone
    FrmPrincipal.GridVendedor.Col = 2
    LblTel.Caption = FrmPrincipal.GridVendedor.Text

End Sub

