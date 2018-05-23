VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmCliente_Alt 
   Caption         =   "Alteração de Cliente"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6930
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
   Icon            =   "FrmCliente_Alt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   6930
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
      Height          =   5655
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   6735
      Begin VB.TextBox TxtTel2 
         Height          =   285
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   10
         ToolTipText     =   "Número do telefone do cliente"
         Top             =   3120
         Width           =   2295
      End
      Begin VB.ComboBox CboSexo 
         Height          =   315
         ItemData        =   "FrmCliente_Alt.frx":0CCA
         Left            =   1200
         List            =   "FrmCliente_Alt.frx":0CCC
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Sexo do cliente"
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox TxtObs 
         Height          =   645
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   14
         ToolTipText     =   "Observação sobre o cliente"
         Top             =   4800
         Width           =   6495
      End
      Begin VB.TextBox TxtCel 
         Height          =   285
         Left            =   4800
         MaxLength       =   15
         TabIndex        =   11
         ToolTipText     =   "Número do celular do cliente"
         Top             =   3120
         Width           =   1815
      End
      Begin VB.ComboBox CboEstado 
         Height          =   315
         ItemData        =   "FrmCliente_Alt.frx":0CCE
         Left            =   4800
         List            =   "FrmCliente_Alt.frx":0CD0
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Estado do cliente"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox TxtCidade 
         Height          =   285
         Left            =   1200
         MaxLength       =   100
         TabIndex        =   4
         ToolTipText     =   "Cidade do cliente"
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox TxtBairro 
         Height          =   285
         Left            =   1200
         MaxLength       =   60
         TabIndex        =   2
         ToolTipText     =   "Bairro do cliente"
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox TxtEndereco 
         Height          =   285
         Left            =   1200
         MaxLength       =   200
         TabIndex        =   1
         ToolTipText     =   "Endereço do cliente"
         Top             =   720
         Width           =   5415
      End
      Begin VB.TextBox TxtNome 
         Height          =   285
         Left            =   1200
         MaxLength       =   200
         TabIndex        =   0
         ToolTipText     =   "Nome do cliente"
         Top             =   240
         Width           =   5415
      End
      Begin VB.TextBox TxtCpf 
         Height          =   285
         Left            =   1200
         MaxLength       =   12
         TabIndex        =   6
         ToolTipText     =   "Cpf do cliente"
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox TxtEmail 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   13
         ToolTipText     =   "Email do cliente"
         Top             =   4080
         Width           =   5415
      End
      Begin VB.TextBox TxtDtNasc 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   7
         Text            =   "__/__/____"
         ToolTipText     =   "Data de nascimento do cliente"
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox TxtCep 
         Height          =   285
         Left            =   4800
         MaxLength       =   9
         TabIndex        =   3
         ToolTipText     =   "Cep do cliente"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox TxtTel1 
         Height          =   285
         Left            =   4800
         MaxLength       =   15
         TabIndex        =   9
         ToolTipText     =   "Número do telefone do cliente"
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox TxtFax 
         Height          =   285
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   12
         ToolTipText     =   "Número do fax do cliente"
         Top             =   3600
         Width           =   2295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCliente_Alt.frx":0CD2
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCliente_Alt.frx":0D34
         TabIndex        =   20
         Top             =   720
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCliente_Alt.frx":0D9E
         TabIndex        =   21
         Top             =   1200
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCliente_Alt.frx":0E04
         TabIndex        =   22
         Top             =   1680
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   3720
         OleObjectBlob   =   "FrmCliente_Alt.frx":0E6A
         TabIndex        =   23
         Top             =   2160
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCliente_Alt.frx":0ECE
         TabIndex        =   24
         Top             =   2160
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   3720
         OleObjectBlob   =   "FrmCliente_Alt.frx":0F2E
         TabIndex        =   25
         Top             =   1200
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   3720
         OleObjectBlob   =   "FrmCliente_Alt.frx":0F8E
         TabIndex        =   26
         Top             =   1680
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   3720
         OleObjectBlob   =   "FrmCliente_Alt.frx":0FF6
         TabIndex        =   27
         Top             =   2640
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   3720
         OleObjectBlob   =   "FrmCliente_Alt.frx":1064
         TabIndex        =   28
         Top             =   3120
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCliente_Alt.frx":10CC
         TabIndex        =   29
         Top             =   4080
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCliente_Alt.frx":1130
         TabIndex        =   30
         Top             =   4560
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCliente_Alt.frx":119E
         TabIndex        =   31
         Top             =   2640
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCliente_Alt.frx":1200
         TabIndex        =   32
         Top             =   3600
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCliente_Alt.frx":1260
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
      Width           =   6735
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   2040
         OleObjectBlob   =   "FrmCliente_Alt.frx":12CE
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
         Left            =   5400
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
         Left            =   4080
         TabIndex        =   15
         ToolTipText     =   "Efetuar a alteração"
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmCliente_Alt"
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
    If TxtNome.Text = "" Or TxtEndereco.Text = "" Or TxtBairro.Text = "" Or TxtCidade.Text = "" Or CboEstado.Text = "" Then
        VPStrBox = MsgBox("Preencha pelo menos os campos principais." & Chr(13) & "(Nome, Endereço, Bairro, Cidade e Estado)", vbInformation, "Pró Vendas 2004 - Informação")
    Else
        Conecta
        
        Dim RecCli As New ADODB.Recordset
        
        StrSql = "SELECT * FROM tb_Cliente where CodCli=" & VGIntCodCli
        RecCli.Open StrSql, vgCon, 1, 3
            
        RecCli("Nome") = TxtNome.Text
        RecCli("Sexo") = CboSexo.Text
        RecCli("Endereco") = TxtEndereco.Text
        RecCli("Bairro") = TxtBairro.Text
        RecCli("Cep") = TxtCep.Text
        RecCli("Cidade") = TxtCidade.Text
        RecCli("Estado") = CboEstado.Text
        RecCli("DtNasc") = FormataDataUS(TxtDtNasc.Text)
        RecCli("Telefone1") = TxtTel1.Text
        RecCli("Telefone2") = TxtTel2.Text
        RecCli("Celular") = TxtCel.Text
        RecCli("Fax") = TxtFax.Text
        RecCli("Cpf") = TxtCpf.Text
        RecCli("Email") = TxtEmail.Text
        RecCli("Obs") = Trim(TxtObs.Text)
        RecCli.Update
            
        VGIntCodCli = 0
        
        Desconecta
        
        VPStrBox = MsgBox("Alteração efetuada.", vbInformation, "Pró Vendas 2004 - Informação")
            
        FrmPrincipal.CmdPesqCli.Value = True
        
        Unload Me
        
        MDIPrincipal.Enabled = True
        MDIPrincipal.WindowState = 2
    End If
End Sub

Private Sub Form_Resize()
  FrmCliente_Alt.Left = (MDIPrincipal.Width / 2) - (FrmCliente_Alt.Width / 2)
  FrmCliente_Alt.Top = (MDIPrincipal.Height / 3) - (FrmCliente_Alt.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 7095
    Width = 7050
    
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    MDIPrincipal.Enabled = False
    
    Call MontaCbos
    
    Conecta
    
    Dim RecCli As New ADODB.Recordset
    Dim VLStrTel1 As String
    Dim VLStrTel2 As String
    
    StrSql = "SELECT * FROM tb_Cliente where CodCli=" & VGIntCodCli
    RecCli.Open StrSql, vgCon, 1, 3
        
    TxtNome.Text = VerificaNulo(RecCli!nome)
    
    If RecCli!sexo <> "" Then
        CboSexo.Text = RecCli!sexo
    End If
    
    TxtEndereco.Text = VerificaNulo(RecCli!endereco)
    TxtBairro.Text = VerificaNulo(RecCli!bairro)
    TxtCep.Text = VerificaNulo(RecCli!cep)
    TxtCidade.Text = VerificaNulo(RecCli!cidade)
    CboEstado.Text = VerificaNulo(RecCli!Estado)
    
    If RecCli!dtnasc <> "" Then
        TxtDtNasc.Text = FormataData(RecCli!dtnasc)
    Else
        TxtDtNasc.Text = "__/__/____"
    End If
    
    TxtTel1.Text = VerificaNulo(RecCli!telefone1)
    TxtTel2.Text = VerificaNulo(RecCli!telefone2)
    TxtFax.Text = VerificaNulo(RecCli!fax)
    TxtCel.Text = VerificaNulo(RecCli!celular)
    TxtCpf.Text = VerificaNulo(RecCli!cpf)
    TxtEmail.Text = VerificaNulo(RecCli!email)
    TxtObs.Text = VerificaNulo(RecCli!obs)
    
    Desconecta
        
End Sub

Sub MontaCbos()
    '===== CboSexo ==============
    CboSexo.AddItem ("")
    CboSexo.AddItem ("Feminino")
    CboSexo.AddItem ("Masculino")
    '============================
    
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

Private Sub TxtCel_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, parênteses, espaço, - e / ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtTel1_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, parênteses, espaço, - e / ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtTel2_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, parênteses, espaço, - e / ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtFax_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, parênteses, espaço, - e / ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtCep_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e - ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtCpf_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e - ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtNasc_GotFocus()
    If TxtDtNasc.Text = "__/__/____" Then
        TxtDtNasc.Text = ""
    End If
    
    TxtDtNasc.SelStart = 0
    TxtDtNasc.SelLength = Len(TxtDtNasc.Text)
End Sub

Private Sub TxtDtNasc_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtNasc_LostFocus()
    
    Dim VLStrData As String
    
    If TxtDtNasc.Text <> "" Then
        VLStrData = VerificaData(TxtDtNasc.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtNasc.SetFocus
        Else
            TxtDtNasc.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtNasc.Text = "__/__/____"
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

Private Sub TxtEndereco_GotFocus()
    TxtEndereco.SelStart = 0
    TxtEndereco.SelLength = Len(TxtEndereco.Text)
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

Private Sub TxtCpf_GotFocus()
    TxtCpf.SelStart = 0
    TxtCpf.SelLength = Len(TxtCpf.Text)
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

Private Sub TxtObs_GotFocus()
    TxtObs.SelStart = 0
    TxtObs.SelLength = Len(TxtObs.Text)
End Sub


