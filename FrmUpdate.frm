VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Atualização do sistema"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
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
      Left            =   960
      TabIndex        =   3
      ToolTipText     =   "Reiniciar atualização do sistema"
      Top             =   1680
      Width           =   2775
   End
   Begin VB.CommandButton CmdOutraVez 
      Caption         =   "&Reiniciar"
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
      Left            =   840
      TabIndex        =   2
      ToolTipText     =   "Reiniciar atualização do sistema"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton CmdAbortar 
      Caption         =   "&Abortar"
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
      Left            =   2760
      TabIndex        =   1
      ToolTipText     =   "Abortar atualização do sistema"
      Top             =   1680
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   360
      OleObjectBlob   =   "FrmUpdate.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblMsg 
      Height          =   1215
      Left            =   120
      OleObjectBlob   =   "FrmUpdate.frx":0234
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "FrmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub CmdAbortar_Click()
'    LblMsg.Caption = "Abortando update do sistema..."
'    CmdAbortar.Enabled = False
'
'    MDIPrincipal.Inet.Cancel
'
'    Unload Me
'    MDIPrincipal.Enabled = True
'    FrmPrincipal.Show
'
'End Sub
'
'Private Sub CmdFechar_Click()
'    Unload Me
'    MDIPrincipal.Enabled = True
'    FrmPrincipal.Show
'End Sub
'
'Private Sub CmdOutraVez_Click()
'    Call MDIPrincipal.FazUpdate
'End Sub
'
'Private Sub Form_Load()
'    Top = 3800
'    Height = 2820
'    Width = 4770
'    Left = 5400
'
'    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
'    Skin1.ApplySkin (Me.hwnd)
'
'    MDIPrincipal.Enabled = False
'    CmdOutraVez.Enabled = False
'    CmdFechar.Visible = False
'End Sub
'
