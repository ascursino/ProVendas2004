VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmCaixa_AReceber_Baixa 
   Caption         =   "Baixa em contas a receber"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6360
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
   Icon            =   "FrmCaixa_AReceber_Baixa.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   6360
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
      Height          =   2055
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   6135
      Begin VB.OptionButton OptRecebChq 
         Caption         =   "Cheque"
         Height          =   255
         Left            =   3240
         TabIndex        =   7
         ToolTipText     =   "Pagamento em cheque"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.OptionButton OptRecebDin 
         Caption         =   "Dinheiro"
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         ToolTipText     =   "Pagamento em dinheiro"
         Top             =   1680
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox TxtDesconto 
         Height          =   285
         Left            =   4680
         MaxLength       =   5
         TabIndex        =   5
         ToolTipText     =   "Desconto"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox TxtJuros 
         Height          =   285
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   4
         ToolTipText     =   "Juros"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox TxtDtReceb 
         Height          =   285
         Left            =   4680
         MaxLength       =   10
         TabIndex        =   3
         Text            =   "__/__/____"
         ToolTipText     =   "Data do recebimento da conta"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TxtValorReceb 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   2
         ToolTipText     =   "Valor recebido da conta"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TxtDtVenc 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4680
         MaxLength       =   10
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "__/__/____"
         ToolTipText     =   "Data de vencimento da conta"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TxtValor 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Valor da conta"
         Top             =   240
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCaixa_AReceber_Baixa.frx":0CCA
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCaixa_AReceber_Baixa.frx":0D2E
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   3360
         OleObjectBlob   =   "FrmCaixa_AReceber_Baixa.frx":0DA4
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   3360
         OleObjectBlob   =   "FrmCaixa_AReceber_Baixa.frx":0E14
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCaixa_AReceber_Baixa.frx":0E82
         TabIndex        =   16
         Top             =   1200
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   2400
         OleObjectBlob   =   "FrmCaixa_AReceber_Baixa.frx":0EE6
         TabIndex        =   17
         Top             =   1200
         Width           =   255
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   3360
         OleObjectBlob   =   "FrmCaixa_AReceber_Baixa.frx":0F40
         TabIndex        =   18
         Top             =   1200
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   5400
         OleObjectBlob   =   "FrmCaixa_AReceber_Baixa.frx":0FAA
         TabIndex        =   19
         Top             =   1200
         Width           =   255
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCaixa_AReceber_Baixa.frx":1004
         TabIndex        =   20
         Top             =   1680
         Width           =   1575
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
      TabIndex        =   10
      Top             =   2160
      Width           =   6135
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   960
         OleObjectBlob   =   "FrmCaixa_AReceber_Baixa.frx":107A
         Top             =   240
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
         Left            =   4800
         TabIndex        =   9
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
         Left            =   3480
         TabIndex        =   8
         ToolTipText     =   "Efetuar o recebimento"
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmCaixa_AReceber_Baixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String

Private Sub CmdFechar_Click()
    Unload Me
    FrmPrincipal.Enabled = True
End Sub

Private Sub CmdOK_Click()
    If TxtValor.Text = "" Or TxtDtVenc.Text = "" Or TxtValorReceb.Text = "" Or TxtDtReceb.Text = "" Then
        VPStrBox = MsgBox("Preencha todos os campos.", vbInformation, "Pró Vendas 2004 - Informação")
    
    ElseIf OptRecebDin.Value = False And OptRecebChq.Value = False Then
        VPStrBox = MsgBox("Escolha a forma de recebimento.", vbInformation, "Pró Vendas 2004 - Informação")
    
    Else
        Conecta
        
        Dim RecBx As New ADODB.Recordset
        Dim RecReceb As New ADODB.Recordset
        Dim RecDescr As New ADODB.Recordset
        Dim RecCxa As New ADODB.Recordset
        Dim TipoReceb As String
        
        If OptRecebDin.Value = True Then
            TipoReceb = "Dinheiro"
        ElseIf OptRecebChq.Value = True Then
            TipoReceb = "Cheque"
        End If
        
        StrSql = "SELECT * FROM tb_ContaReceber_Recebido"
        RecBx.Open StrSql, vgCon, 1, 3
        
        RecBx.AddNew
        RecBx("CodCReceb") = VGIntCodReceber
        RecBx("DtReceb") = FormataDataUS(TxtDtReceb.Text)
        RecBx("Juros") = TxtJuros.Text
        RecBx("Desconto") = TxtDesconto.Text
        RecBx("ValorReceb") = Mid(TxtValorReceb.Text, 4)
        RecBx.Update
        
        'atualiza informações de recebimento na tabela de contas a receber
        StrSql = "SELECT * FROM tb_ContaReceber where CodCReceb=" & VGIntCodReceber
        RecReceb.Open StrSql, vgCon, 1, 3
        
        RecReceb("Recebido") = "sim"
        RecReceb.Update
        
        VPStrResponse = MsgBox("Envia dados para o lançamento de caixa?", vbYesNo, "Pró Vendas 2004 - Informação")
        
        If VPStrResponse = vbYes Then
            'envia dados para o lançamento de caixa
            StrSql = "SELECT * FROM tb_Caixa"
            RecCxa.Open StrSql, vgCon, 1, 3
            
            StrSql = "SELECT Descricao FROM tb_ContaReceber where CodCReceb=" & VGIntCodReceber
            RecDescr.Open StrSql, vgCon, 1, 3
        
            RecCxa.AddNew
            RecCxa("CodVenda") = 0
            RecCxa("DtMov") = FormataDataUS(TxtDtReceb.Text)
            RecCxa("TipoMov") = "Baixa no recebimento de contas"
            RecCxa("Valor") = Mid(TxtValorReceb.Text, 4)
            RecCxa("TipoValor") = "credito"
            RecCxa("Descricao") = RecDescr!Descricao
            RecCxa("TipoPagto") = TipoReceb
            RecCxa.Update
        End If
        
        Desconecta
        
        VPStrBox = MsgBox("Baixa efetuada.", vbInformation, "Pró Vendas 2004 - Informação")
        
        'atualiza a grid de contas a receber
        FrmPrincipal.CmdPesqAReceber.Value = True
        
        Unload Me
        
        FrmPrincipal.Enabled = True
    End If
End Sub

Private Sub Form_Resize()
  FrmCaixa_AReceber_Baixa.Left = (MDIPrincipal.Width / 2) - (FrmCaixa_AReceber_Baixa.Width / 2)
  FrmCaixa_AReceber_Baixa.Top = (MDIPrincipal.Height / 3) - (FrmCaixa_AReceber_Baixa.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 3510
    Width = 6480
    
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    FrmPrincipal.Enabled = False
    
    Conecta
    
    Dim RecBx As New ADODB.Recordset
    
    StrSql = "SELECT * FROM tb_ContaReceber WHERE CodCReceb=" & VGIntCodReceber
    RecBx.Open StrSql, vgCon, 1, 3
    
    TxtValor.Text = FormataMoeda(VerificaNulo(RecBx!valor))
    TxtDtVenc.Text = FormataData(RecBx!vencimento)
    TxtValorReceb.Text = FormataMoeda(VerificaNulo(RecBx!valor))
    TxtDtReceb.Text = FormataData(Date)
    
    If (IsNull(RecBx!Numbanco) = True Or RecBx!Numbanco = "0") And (IsNull(RecBx!Numcheque) = True Or RecBx!Numcheque = "") Then
        OptRecebChq.Value = False
        OptRecebDin.Value = True
    Else
        OptRecebChq.Value = True
        OptRecebDin.Value = False
    End If
    
    Desconecta
    
End Sub

Private Sub TxtDesconto_GotFocus()
    TxtDesconto.SelStart = 0
    TxtDesconto.SelLength = Len(TxtDesconto.Text)
End Sub

Private Sub TxtDtReceb_GotFocus()
    If TxtDtReceb.Text = "__/__/____" Then
        TxtDtReceb.Text = ""
    End If
    TxtDtReceb.SelStart = 0
    TxtDtReceb.SelLength = Len(TxtDtReceb.Text)
End Sub

Private Sub TxtDtReceb_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e barra ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtReceb_LostFocus()
    Dim VLStrData As String
    
    If TxtDtReceb.Text <> "" Then
        VLStrData = VerificaData(TxtDtReceb.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtReceb.SetFocus
        Else
            TxtDtReceb.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtReceb.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtVenc_GotFocus()
    If TxtDtVenc.Text = "__/__/____" Then
        TxtDtVenc.Text = ""
    End If
    TxtDtVenc.SelStart = 0
    TxtDtVenc.SelLength = Len(TxtDtVenc.Text)
End Sub

Private Sub TxtDtVenc_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e barra ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtVenc_LostFocus()
    Dim VLStrData As String
    
    If TxtDtVenc.Text <> "" Then
        VLStrData = VerificaData(TxtDtVenc.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtVenc.SetFocus
        Else
            TxtDtVenc.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtVenc.Text = "__/__/____"
    End If
End Sub

Private Sub TxtJuros_GotFocus()
    TxtJuros.SelStart = 0
    TxtJuros.SelLength = Len(TxtJuros.Text)
End Sub

Private Sub TxtValor_GotFocus()
    TxtValor.SelStart = 0
    TxtValor.SelLength = Len(TxtValor.Text)
End Sub

Private Sub TxtValor_LostFocus()
    TxtValor.Text = FormataMoeda(TxtValor.Text)
End Sub

Private Sub TxtValorReceb_GotFocus()
    TxtValorReceb.SelStart = 0
    TxtValorReceb.SelLength = Len(TxtValorReceb.Text)
End Sub

Private Sub TxtValorReceb_LostFocus()
    TxtValorReceb.Text = FormataMoeda(TxtValorReceb.Text)
End Sub
