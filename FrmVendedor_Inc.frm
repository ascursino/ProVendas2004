VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmVendedor_Inc 
   Caption         =   "Inclus�o de Vendedor"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4425
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
   Icon            =   "FrmVendedor_Inc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   4425
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
      TabIndex        =   5
      Top             =   1440
      Width           =   4215
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   480
         OleObjectBlob   =   "FrmVendedor_Inc.frx":0CCA
         Top             =   120
      End
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
         Left            =   2880
         TabIndex        =   3
         ToolTipText     =   "Fechar"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "&Ok"
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
         Left            =   1560
         TabIndex        =   2
         ToolTipText     =   "Efetuar inclus�o"
         Top             =   240
         Width           =   1095
      End
   End
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
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4215
      Begin VB.TextBox TxtTel 
         Height          =   285
         Left            =   1080
         MaxLength       =   14
         TabIndex        =   1
         ToolTipText     =   "N�mero do telefone do vendedor"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TxtVend 
         Height          =   285
         Left            =   1080
         MaxLength       =   200
         TabIndex        =   0
         ToolTipText     =   "Nome do vendedor"
         Top             =   240
         Width           =   3015
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVendedor_Inc.frx":0EFE
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVendedor_Inc.frx":0F68
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmVendedor_Inc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String

Private Sub CmdFechar_Click()
    Unload Me
    MDIPrincipal.Enabled = True
    FrmPrincipal.CmdPesqVendedor.Enabled = True
End Sub

Private Sub CmdOK_Click()
    If TxtVend.Text = "" Then
        VPStrBox = MsgBox("Campo 'Vendedor' n�o pode ser em branco.", vbInformation, "Pr� Vendas 2004 - Informa��o")
    Else
        Conecta
        
        Dim RecVend As New ADODB.Recordset
        
        StrSql = "Select * From tb_Vendedor"
        RecVend.Open StrSql, vgCon, 1, 3
        
        RecVend.AddNew
        RecVend("Nome") = TxtVend.Text
        RecVend("Telefone") = TxtTel.Text
        RecVend.Update
        
        Desconecta
        
        FrmPrincipal.CmdPesqVendedor.Value = True
        
        VPStrBox = MsgBox("Vendedor cadastrado.", vbInformation, "Pr� Vendas 2004 - Informa��o")
        
        TxtVend.Text = ""
        TxtTel.Text = ""
        
        TxtVend.SetFocus
    End If
    
End Sub

Private Sub Form_Resize()
  FrmVendedor_Inc.Left = (MDIPrincipal.Width / 2) - (FrmVendedor_Inc.Width / 2)
  FrmVendedor_Inc.Top = (MDIPrincipal.Height / 3) - (FrmVendedor_Inc.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 2760
    Width = 4545
    
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    MDIPrincipal.Enabled = False
End Sub

Private Sub TxtTel_GotFocus()
    TxtTel.SelStart = 0
    TxtTel.SelLength = Len(TxtTel.Text)
End Sub

Private Sub TxtTel_KeyPress(KeyAscii As Integer)
    '=== S� aceita n�meros, par�nteses, espa�o e - ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtVend_GotFocus()
    TxtVend.SelStart = 0
    TxtVend.SelLength = Len(TxtVend.Text)
End Sub
