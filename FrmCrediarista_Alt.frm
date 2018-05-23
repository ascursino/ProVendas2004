VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmCrediarista_Alt 
   Caption         =   "Alteração de Crediarista"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6960
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
   Icon            =   "FrmCrediarista_Alt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   6960
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
      Height          =   4695
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   6735
      Begin VB.TextBox TxtObs 
         Height          =   645
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "Observação sobre o crediarista"
         Top             =   3840
         Width           =   6495
      End
      Begin VB.TextBox TxtTel2 
         Height          =   285
         Left            =   5160
         MaxLength       =   15
         TabIndex        =   9
         ToolTipText     =   "Número do telefone do crediarista"
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox TxtTel1 
         Height          =   285
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   8
         ToolTipText     =   "Número do telefone do crediarista"
         Top             =   2640
         Width           =   1455
      End
      Begin VB.ComboBox CboEstado 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Estado do crediarista"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox TxtCidade 
         Height          =   285
         Left            =   1200
         MaxLength       =   100
         TabIndex        =   4
         ToolTipText     =   "Cidade do crediarista"
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox TxtBairro 
         Height          =   285
         Left            =   1200
         MaxLength       =   60
         TabIndex        =   2
         ToolTipText     =   "Bairro do crediarista"
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox TxtEndereco 
         Height          =   285
         Left            =   1200
         MaxLength       =   200
         TabIndex        =   1
         ToolTipText     =   "Endereço do crediarista"
         Top             =   720
         Width           =   5415
      End
      Begin VB.TextBox TxtNome 
         Height          =   285
         Left            =   1200
         MaxLength       =   200
         TabIndex        =   0
         ToolTipText     =   "Nome do crediarista"
         Top             =   240
         Width           =   5415
      End
      Begin VB.TextBox TxtEmail 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   10
         ToolTipText     =   "Email do crediarista"
         Top             =   3120
         Width           =   5415
      End
      Begin VB.TextBox TxtCpf 
         Height          =   285
         Left            =   1200
         MaxLength       =   14
         TabIndex        =   6
         ToolTipText     =   "Cpf do crediarista"
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox TxtCep 
         Height          =   285
         Left            =   5160
         MaxLength       =   9
         TabIndex        =   3
         ToolTipText     =   "Cep do crediarista"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox TxtDtNasc 
         Height          =   285
         Left            =   5160
         MaxLength       =   10
         TabIndex        =   7
         Text            =   "__/__/____"
         ToolTipText     =   "Data do nascimento do crediarista"
         Top             =   2160
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediarista_Alt.frx":0CCA
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediarista_Alt.frx":0D2C
         TabIndex        =   17
         Top             =   720
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediarista_Alt.frx":0D96
         TabIndex        =   18
         Top             =   1200
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediarista_Alt.frx":0DFC
         TabIndex        =   19
         Top             =   1680
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   4080
         OleObjectBlob   =   "FrmCrediarista_Alt.frx":0E62
         TabIndex        =   20
         Top             =   2160
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediarista_Alt.frx":0EC6
         TabIndex        =   21
         Top             =   2160
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   4080
         OleObjectBlob   =   "FrmCrediarista_Alt.frx":0F26
         TabIndex        =   22
         Top             =   1200
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   4080
         OleObjectBlob   =   "FrmCrediarista_Alt.frx":0F86
         TabIndex        =   23
         Top             =   1680
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediarista_Alt.frx":0FEC
         TabIndex        =   24
         Top             =   2640
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   4080
         OleObjectBlob   =   "FrmCrediarista_Alt.frx":105A
         TabIndex        =   25
         Top             =   2640
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediarista_Alt.frx":10C8
         TabIndex        =   26
         Top             =   3120
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediarista_Alt.frx":112C
         TabIndex        =   27
         Top             =   3600
         Width           =   1335
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
      TabIndex        =   14
      Top             =   4800
      Width           =   6735
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   2040
         OleObjectBlob   =   "FrmCrediarista_Alt.frx":119A
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
         TabIndex        =   13
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
         TabIndex        =   12
         ToolTipText     =   "Efetuar alteração"
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmCrediarista_Alt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String

Private Sub CmdFechar_Click()
    VGIntCodCredsta = 0
    
    Unload Me
    
    MDIPrincipal.Enabled = True
    MDIPrincipal.WindowState = 2
End Sub

Private Sub CmdOK_Click()
    If TxtNome.Text = "" Or TxtEndereco.Text = "" Or TxtBairro.Text = "" Or TxtCidade.Text = "" Or CboEstado.Text = "" Then
        VPStrBox = MsgBox("Preencha pelo menos os campos principais." & Chr(13) & "(Nome, Endereço, Bairro, Cidade e Estado)", vbInformation, "Pró Vendas 2004 - Informação")
    Else
        Conecta
        
        Dim RecCredsta As New ADODB.Recordset
        
        StrSql = "SELECT * FROM tb_Crediarista where CodCredsta=" & VGIntCodCredsta
        RecCredsta.Open StrSql, vgCon, 1, 3
            
        RecCredsta("Nome") = TxtNome.Text
        RecCredsta("Endereco") = TxtEndereco.Text
        RecCredsta("Bairro") = TxtBairro.Text
        RecCredsta("Cep") = TxtCep.Text
        RecCredsta("Cidade") = TxtCidade.Text
        RecCredsta("Estado") = CboEstado.Text
        RecCredsta("DtNasc") = FormataDataUS(TxtDtNasc.Text)
        RecCredsta("Telefone1") = TxtTel1.Text
        RecCredsta("Telefone2") = TxtTel2.Text
        RecCredsta("Cpf") = TxtCpf.Text
        RecCredsta("Email") = TxtEmail.Text
        RecCredsta("Obs") = Trim(TxtObs.Text)
        RecCredsta.Update
            
        VGIntCodCredsta = 0
        
        Desconecta
        
        VPStrBox = MsgBox("Alteração efetuada.", vbInformation, "Pró Vendas 2004 - Informação")
            
        FrmPrincipal.CmdPesqCredsta.Value = True
        
        Unload Me
        
        MDIPrincipal.Enabled = True
        MDIPrincipal.WindowState = 2
    End If
End Sub

Private Sub Form_Resize()
  FrmCrediarista_Alt.Left = (MDIPrincipal.Width / 2) - (FrmCrediarista_Alt.Width / 2)
  FrmCrediarista_Alt.Top = (MDIPrincipal.Height / 3) - (FrmCrediarista_Alt.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 6135
    Width = 7080
    
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    Call MontaCbos
    
    Conecta
    
    Dim RecCredsta As New ADODB.Recordset
    Dim VLIntCodCredsta As Integer
    
    If VGIntCodCredsta = 0 Then
        VLIntCodCredsta = VGIntCodCredstaVenda
    Else
        VLIntCodCredsta = VGIntCodCredsta
    End If
    
    StrSql = "SELECT * FROM tb_Crediarista where CodCredsta=" & VLIntCodCredsta
    RecCredsta.Open StrSql, vgCon, 1, 3
    
    If Not RecCredsta.EOF Then
        TxtNome.Text = VerificaNulo(RecCredsta!nome)
        TxtEndereco.Text = VerificaNulo(RecCredsta!endereco)
        TxtBairro.Text = VerificaNulo(RecCredsta!bairro)
        TxtCep.Text = VerificaNulo(RecCredsta!cep)
        TxtCidade.Text = VerificaNulo(RecCredsta!cidade)
        
        If RecCredsta!Estado <> "" And IsNull(RecCredsta!Estado) = False Then
            CboEstado.Text = RecCredsta!Estado
        End If
        
        If RecCredsta!dtnasc <> "" And IsNull(RecCredsta!dtnasc) = False Then
            TxtDtNasc.Text = FormataData(RecCredsta!dtnasc)
        Else
            TxtDtNasc.Text = "__/__/____"
        End If
        TxtTel1.Text = VerificaNulo(RecCredsta!telefone1)
        TxtTel2.Text = VerificaNulo(RecCredsta!telefone2)
        TxtCpf.Text = VerificaNulo(RecCredsta!cpf)
        TxtEmail.Text = VerificaNulo(RecCredsta!email)
        TxtObs.Text = VerificaNulo(RecCredsta!obs)
    End If
    Desconecta
    
    MDIPrincipal.Enabled = False
    
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

Private Sub TxtEmail_GotFocus()
    TxtEmail.SelStart = 0
    TxtEmail.SelLength = Len(TxtEmail.Text)
End Sub

Private Sub TxtEndereco_GotFocus()
    TxtEndereco.SelStart = 0
    TxtEndereco.SelLength = Len(TxtEndereco.Text)
End Sub

Private Sub TxtNome_GotFocus()
    TxtNome.SelStart = 0
    TxtNome.SelLength = Len(TxtNome.Text)
End Sub

Private Sub TxtObs_GotFocus()
    TxtObs.SelStart = 0
    TxtObs.SelLength = Len(TxtObs.Text)
End Sub

Private Sub TxtTel1_GotFocus()
    TxtTel1.SelStart = 0
    TxtTel1.SelLength = Len(TxtTel1.Text)
End Sub

Private Sub TxtTel1_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, parênteses, espaço, - e / ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtTel2_GotFocus()
    TxtTel2.SelStart = 0
    TxtTel2.SelLength = Len(TxtTel2.Text)
End Sub

Private Sub TxtTel2_KeyPress(KeyAscii As Integer)
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
    '=== Só aceita números e / ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
    
    If TxtDtNasc.Text = "__/__/____" Then
        TxtDtNasc.Text = ""
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
