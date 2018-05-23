VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmFornecedor_Alt 
   Caption         =   "Alteração de Fornecedor"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
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
   Icon            =   "FrmFornecedor_Alt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   7095
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   6855
      Begin VB.TextBox TxtContato 
         Height          =   285
         Left            =   4920
         MaxLength       =   200
         TabIndex        =   12
         ToolTipText     =   "Nome da pessoa de contato"
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox TxtObs 
         Height          =   645
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   14
         ToolTipText     =   "Observação sobre o fornecedor"
         Top             =   4680
         Width           =   6615
      End
      Begin VB.TextBox TxtCel 
         Height          =   285
         Left            =   4920
         MaxLength       =   15
         TabIndex        =   10
         ToolTipText     =   "Número do celular do fornecedor"
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox TxtTel1 
         Height          =   285
         Left            =   4920
         MaxLength       =   15
         TabIndex        =   8
         ToolTipText     =   "Número do telefone do fornecedor"
         Top             =   2640
         Width           =   1815
      End
      Begin VB.ComboBox CboEstado 
         Height          =   315
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Estado do fornecedor"
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox TxtCidade 
         Height          =   285
         Left            =   1200
         MaxLength       =   100
         TabIndex        =   5
         ToolTipText     =   "Cidade do fornecedor"
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox TxtBairro 
         Height          =   285
         Left            =   1200
         MaxLength       =   60
         TabIndex        =   3
         ToolTipText     =   "Bairro do fornecedor"
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox TxtEndereco 
         Height          =   285
         Left            =   1200
         MaxLength       =   200
         TabIndex        =   2
         ToolTipText     =   "Endereço do fornecedor"
         Top             =   1200
         Width           =   5535
      End
      Begin VB.TextBox TxtNome 
         Height          =   285
         Left            =   1200
         MaxLength       =   200
         TabIndex        =   1
         ToolTipText     =   "Nome do fornecedor"
         Top             =   720
         Width           =   5535
      End
      Begin VB.TextBox TxtCnpj 
         Height          =   285
         Left            =   1200
         MaxLength       =   18
         TabIndex        =   7
         ToolTipText     =   "Cnpj do fornecedor"
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox TxtEmail 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   13
         ToolTipText     =   "Email do fornecedor"
         Top             =   4080
         Width           =   5535
      End
      Begin VB.TextBox TxtCep 
         Height          =   285
         Left            =   4920
         MaxLength       =   9
         TabIndex        =   4
         ToolTipText     =   "Cep do fornecedor"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox TxtTipoForn 
         Height          =   285
         Left            =   1680
         MaxLength       =   100
         TabIndex        =   0
         ToolTipText     =   "Tipo de produto"
         Top             =   240
         Width           =   5055
      End
      Begin VB.TextBox TxtFax 
         Height          =   285
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   11
         ToolTipText     =   "Número do fax do fornecedor"
         Top             =   3600
         Width           =   2295
      End
      Begin VB.TextBox TxtTel2 
         Height          =   285
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   9
         ToolTipText     =   "Número do telefone do fornecedor"
         Top             =   3120
         Width           =   2295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmFornecedor_Alt.frx":0CCA
         TabIndex        =   19
         Top             =   720
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmFornecedor_Alt.frx":0D38
         TabIndex        =   20
         Top             =   1200
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmFornecedor_Alt.frx":0DA2
         TabIndex        =   21
         Top             =   1680
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmFornecedor_Alt.frx":0E08
         TabIndex        =   22
         Top             =   2160
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmFornecedor_Alt.frx":0E6E
         TabIndex        =   23
         Top             =   2640
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmFornecedor_Alt.frx":0ED0
         TabIndex        =   24
         Top             =   4080
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   3840
         OleObjectBlob   =   "FrmFornecedor_Alt.frx":0F34
         TabIndex        =   25
         Top             =   3600
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   3840
         OleObjectBlob   =   "FrmFornecedor_Alt.frx":0F9C
         TabIndex        =   26
         Top             =   1680
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   3840
         OleObjectBlob   =   "FrmFornecedor_Alt.frx":0FFC
         TabIndex        =   27
         Top             =   2160
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   3840
         OleObjectBlob   =   "FrmFornecedor_Alt.frx":1062
         TabIndex        =   28
         Top             =   2640
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   3840
         OleObjectBlob   =   "FrmFornecedor_Alt.frx":10D0
         TabIndex        =   29
         Top             =   3120
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmFornecedor_Alt.frx":1138
         TabIndex        =   30
         Top             =   4440
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmFornecedor_Alt.frx":11A6
         TabIndex        =   31
         Top             =   240
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmFornecedor_Alt.frx":121E
         TabIndex        =   32
         Top             =   3600
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmFornecedor_Alt.frx":127E
         TabIndex        =   33
         Top             =   3120
         Width           =   975
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
      TabIndex        =   17
      Top             =   5760
      Width           =   6855
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   720
         OleObjectBlob   =   "FrmFornecedor_Alt.frx":12EC
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
         Left            =   5520
         TabIndex        =   16
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
         Left            =   4200
         TabIndex        =   15
         ToolTipText     =   "Efetuar alteração"
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmFornecedor_Alt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String

Private Sub CmdFechar_Click()
    Unload Me
    
    MDIPrincipal.Enabled = True
    MDIPrincipal.WindowState = 2
End Sub

Private Sub CmdOK_Click()
    If TxtTipoForn.Text = "" Or TxtNome.Text = "" Or TxtEndereco.Text = "" Or TxtBairro.Text = "" Or TxtCidade.Text = "" Or CboEstado.Text = "" Then
        VPStrBox = MsgBox("Preencha pelo menos os campos principais." & Chr(13) & "(Tipo de produto, Fornecedor, Endereço, Bairro, Cidade e Estado)", vbInformation, "Pró Vendas 2004 - Informação")
    Else
        Conecta
        
        Dim RecForn As New ADODB.Recordset
        
        StrSql = "SELECT * FROM tb_Fornecedor where CodForn=" & VGIntCodForn
        RecForn.Open StrSql, vgCon, 1, 3
            
        RecForn("Tipo") = TxtTipoForn.Text
        RecForn("Nome") = TxtNome.Text
        RecForn("Endereco") = TxtEndereco.Text
        RecForn("Bairro") = TxtBairro.Text
        RecForn("Cep") = TxtCep.Text
        RecForn("Cidade") = TxtCidade.Text
        RecForn("Estado") = CboEstado.Text
        RecForn("CNPJ") = TxtCnpj.Text
        RecForn("Email") = TxtEmail.Text
        RecForn("Contato") = TxtContato.Text
        RecForn("Telefone1") = TxtTel1.Text
        RecForn("Telefone2") = TxtTel2.Text
        RecForn("Celular") = TxtCel.Text
        RecForn("Fax") = TxtFax.Text
        RecForn("Obs") = Trim(TxtObs.Text)
        RecForn.Update
            
        VGIntCodForn = 0
        
        Desconecta
        
        VPStrBox = MsgBox("Alteração efetuada.", vbInformation, "Pró Vendas 2004 - Informação")
            
        FrmPrincipal.CmdPesqForn.Value = True
        
        Unload Me
        
        MDIPrincipal.Enabled = True
        MDIPrincipal.WindowState = 2
    End If
End Sub

Private Sub Form_Resize()
  FrmFornecedor_Alt.Left = (MDIPrincipal.Width / 2) - (FrmFornecedor_Alt.Width / 2)
  FrmFornecedor_Alt.Top = (MDIPrincipal.Height / 3) - (FrmFornecedor_Alt.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 7095
    Width = 7215
    
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    MDIPrincipal.Enabled = False
    
    Call MontaCbos
    
    Conecta
    
    Dim RecForn As New ADODB.Recordset
    
    StrSql = "SELECT * FROM tb_Fornecedor where CodForn=" & VGIntCodForn
    RecForn.Open StrSql, vgCon, 1, 3
        
    TxtTipoForn.Text = VerificaNulo(RecForn!tipo)
    TxtNome.Text = VerificaNulo(RecForn!nome)
    TxtEndereco.Text = VerificaNulo(RecForn!endereco)
    TxtBairro.Text = VerificaNulo(RecForn!bairro)
    TxtCep.Text = VerificaNulo(RecForn!cep)
    TxtCidade.Text = VerificaNulo(RecForn!cidade)
    CboEstado.Text = VerificaNulo(RecForn!Estado)
    TxtCnpj.Text = VerificaNulo(RecForn!cnpj)
    TxtEmail.Text = VerificaNulo(RecForn!email)
    TxtContato.Text = VerificaNulo(RecForn!contato)
    TxtTel1.Text = VerificaNulo(RecForn!telefone1)
    TxtTel2.Text = VerificaNulo(RecForn!telefone2)
    TxtCel.Text = VerificaNulo(RecForn!celular)
    TxtFax.Text = VerificaNulo(RecForn!fax)
    TxtObs.Text = VerificaNulo(RecForn!obs)
    
    Desconecta
    
End Sub

Sub MontaCbos()
    '===== CboEstado ============
    CboEstado.AddItem ("")
    CboEstado.AddItem ("AC")
    CboEstado.AddItem ("AL")
    CboEstado.AddItem ("AM")
    CboEstado.AddItem ("AP")
    CboEstado.AddItem ("BA")
    CboEstado.AddItem ("CE")
    CboEstado.AddItem ("DF")
    CboEstado.AddItem ("ES")
    CboEstado.AddItem ("GO")
    CboEstado.AddItem ("MA")
    CboEstado.AddItem ("MG")
    CboEstado.AddItem ("MS")
    CboEstado.AddItem ("MT")
    CboEstado.AddItem ("PA")
    CboEstado.AddItem ("PB")
    CboEstado.AddItem ("PE")
    CboEstado.AddItem ("PI")
    CboEstado.AddItem ("PR")
    CboEstado.AddItem ("RJ")
    CboEstado.AddItem ("RN")
    CboEstado.AddItem ("RO")
    CboEstado.AddItem ("RR")
    CboEstado.AddItem ("RS")
    CboEstado.AddItem ("SC")
    CboEstado.AddItem ("SE")
    CboEstado.AddItem ("SP")
    CboEstado.AddItem ("TO")
    '============================
End Sub

Private Sub TxtTel1_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, parênteses, espaço e - ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtTel2_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, parênteses, espaço e - ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtCel_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, parênteses, espaço e - ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtFax_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, parênteses, espaço e - ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtCep_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtCnpj_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtEmail_LostFocus()
    If TxtEmail.Text <> "" Then
        If InStr(TxtEmail.Text, "@") = 0 Then
            VPStrBox = MsgBox("Formato do email está incorreto.", vbCritical, "Pró Vendas 2004 - Erro")
            TxtEmail.SetFocus
        End If
    End If
End Sub

Private Sub TxtNome_GotFocus()
    TxtNome.SelStart = 0
    TxtNome.SelLength = Len(TxtNome.Text)
End Sub

Private Sub TxtTipoForn_GotFocus()
    TxtTipoForn.SelStart = 0
    TxtTipoForn.SelLength = Len(TxtTipoForn.Text)
End Sub

Private Sub TxtBairro_GotFocus()
    TxtBairro.SelStart = 0
    TxtBairro.SelLength = Len(TxtBairro.Text)
End Sub

Private Sub TxtCep_GotFocus()
    TxtCep.SelStart = 0
    TxtCep.SelLength = Len(TxtCep.Text)
End Sub

Private Sub TxtCidade_GotFocus()
    TxtCidade.SelStart = 0
    TxtCidade.SelLength = Len(TxtCidade.Text)
End Sub

Private Sub TxtEndereco_GotFocus()
    TxtEndereco.SelStart = 0
    TxtEndereco.SelLength = Len(TxtEndereco.Text)
End Sub

Private Sub TxtCnpj_GotFocus()
    TxtCnpj.SelStart = 0
    TxtCnpj.SelLength = Len(TxtCnpj.Text)
End Sub

Private Sub TxtTel1_GotFocus()
    TxtTel1.SelStart = 0
    TxtTel1.SelLength = Len(TxtTel1.Text)
End Sub

Private Sub TxtTel2_GotFocus()
    TxtTel2.SelStart = 0
    TxtTel2.SelLength = Len(TxtTel2.Text)
End Sub

Private Sub TxtCel_GotFocus()
    TxtCel.SelStart = 0
    TxtCel.SelLength = Len(TxtCel.Text)
End Sub

Private Sub TxtFax_GotFocus()
    TxtFax.SelStart = 0
    TxtFax.SelLength = Len(TxtFax.Text)
End Sub

Private Sub TxtEmail_GotFocus()
    TxtEmail.SelStart = 0
    TxtEmail.SelLength = Len(TxtEmail.Text)
End Sub

Private Sub TxtContato_GotFocus()
    TxtContato.SelStart = 0
    TxtContato.SelLength = Len(TxtContato.Text)
End Sub

Private Sub TxtObs_GotFocus()
    TxtObs.SelStart = 0
    TxtObs.SelLength = Len(TxtObs.Text)
End Sub

