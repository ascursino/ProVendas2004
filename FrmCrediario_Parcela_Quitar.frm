VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmCrediario_Parcela_Quitar 
   Caption         =   "Quita��o de parcelas"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
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
   Icon            =   "FrmCrediario_Parcela_Quitar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   5865
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
      Height          =   3495
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   5655
      Begin VB.TextBox TxtJuros 
         Height          =   285
         Left            =   4080
         MaxLength       =   5
         TabIndex        =   1
         ToolTipText     =   "Juros"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox TxtDesc 
         Height          =   285
         Left            =   4080
         MaxLength       =   5
         TabIndex        =   3
         ToolTipText     =   "Desconto"
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton OptPagtoDin 
         Caption         =   "Dinheiro"
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         ToolTipText     =   "Pagamento em dinheiro"
         Top             =   2160
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton OptPagtoChq 
         Caption         =   "Cheque"
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         ToolTipText     =   "Pagamento em cheque"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Frame FraPagtoChq 
         Height          =   855
         Left            =   1680
         TabIndex        =   15
         Top             =   2520
         Visible         =   0   'False
         Width           =   2775
         Begin VB.TextBox TxtDigito 
            Height          =   285
            Left            =   2400
            MaxLength       =   2
            TabIndex        =   10
            ToolTipText     =   "D�gito do n�mero do cheque"
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox TxtCheque 
            Height          =   285
            Left            =   1080
            MaxLength       =   17
            TabIndex        =   9
            ToolTipText     =   "N�mero do cheque"
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox TxtBanco 
            Height          =   285
            Left            =   1080
            MaxLength       =   3
            TabIndex        =   8
            ToolTipText     =   "N�mero do banco do cheque"
            Top             =   120
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmCrediario_Parcela_Quitar.frx":0CCA
            TabIndex        =   16
            Top             =   120
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmCrediario_Parcela_Quitar.frx":0D34
            TabIndex        =   17
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.TextBox TxtDtQuit 
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "__/__/____"
         ToolTipText     =   "Data do pagamento da parcela"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox TxtParcPagar 
         Height          =   285
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   4
         ToolTipText     =   "Valor pago da parcela"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox TxtDtVenc 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "__/__/____"
         ToolTipText     =   "Data do vencimento da parcela"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TxtValorParc 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   0
         ToolTipText     =   "Valor da parcela"
         Top             =   240
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediario_Parcela_Quitar.frx":0DA0
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   3120
         OleObjectBlob   =   "FrmCrediario_Parcela_Quitar.frx":0E1A
         TabIndex        =   19
         Top             =   240
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediario_Parcela_Quitar.frx":0E7E
         TabIndex        =   20
         Top             =   1200
         Width           =   2295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediario_Parcela_Quitar.frx":0F02
         TabIndex        =   21
         Top             =   1680
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediario_Parcela_Quitar.frx":0F70
         TabIndex        =   22
         Top             =   2160
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediario_Parcela_Quitar.frx":0FE2
         TabIndex        =   23
         Top             =   720
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   3120
         OleObjectBlob   =   "FrmCrediario_Parcela_Quitar.frx":1050
         TabIndex        =   24
         Top             =   720
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   5160
         OleObjectBlob   =   "FrmCrediario_Parcela_Quitar.frx":10BA
         TabIndex        =   25
         Top             =   720
         Width           =   255
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   5160
         OleObjectBlob   =   "FrmCrediario_Parcela_Quitar.frx":1114
         TabIndex        =   26
         Top             =   240
         Width           =   255
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
      TabIndex        =   13
      Top             =   3600
      Width           =   5655
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   960
         OleObjectBlob   =   "FrmCrediario_Parcela_Quitar.frx":116E
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
         Left            =   4320
         TabIndex        =   12
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
         Left            =   3000
         TabIndex        =   11
         ToolTipText     =   "Efetua a quita��o"
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmCrediario_Parcela_Quitar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String

Private Sub CmdFechar_Click()
    Unload Me
    
    FrmResumo_Crediario.Enabled = True
    MDIPrincipal.Enabled = True
End Sub

Private Sub CmdOK_Click()
    If TxtValorParc.Text = "" Or TxtDtVenc.Text = "" Or TxtDtVenc.Text = "__/__/____" Or TxtParcPagar.Text = "" Or TxtDtQuit.Text = "" Or TxtDtQuit.Text = "__/__/____" Then
        VPStrBox = MsgBox("Preencha todos os campos.", vbInformation, "Pr� Vendas 2004 - Informa��o")
    
    ElseIf OptPagtoChq.Value = True And (TxtBanco.Text = "" Or TxtCheque.Text = "" Or TxtDigito.Text = "") Then
        VPStrBox = MsgBox("Preencha os dados do cheque.", vbInformation, "Pr� Vendas 2004 - Informa��o")
    
    Else
        Conecta
        
        Dim RecQuit As New ADODB.Recordset
        Dim RecCxa As New ADODB.Recordset
        Dim RecParc As New ADODB.Recordset
        Dim RecCli As New ADODB.Recordset
        Dim TipoPagto As String
        
        'grava informa��es na tabela de quita��o de parcelas
        If OptPagtoDin.Value = True Then
            TipoPagto = "Dinheiro"
        ElseIf OptPagtoChq.Value = True Then
            TipoPagto = "Cheque"
        End If
        
        StrSql = "SELECT * FROM tb_Crediario_Parcela_Quitacao"
        RecQuit.Open StrSql, vgCon, 1, 3
        
        RecQuit.AddNew
        RecQuit("CodParc") = VGIntCodParc
        RecQuit("CodCred") = VGIntCodCred
        RecQuit("DtPagto") = FormataDataUS(TxtDtQuit.Text)
        RecQuit("Juros") = TxtJuros.Text
        RecQuit("Desconto") = TxtDesc.Text
        RecQuit("ValorPago") = CCur(TxtParcPagar.Text)
        RecQuit("TipoPagto") = TipoPagto
        
        If TipoPagto = "Cheque" Then
            RecQuit("NumBanco") = TxtBanco.Text
            RecQuit("NumCheque") = TxtCheque.Text & "-" & TxtDigito.Text
        Else
            RecQuit("NumBanco") = 0
            RecQuit("NumCheque") = ""
        End If
        
        RecQuit.Update
        
        'atualiza informa��es de quita��o na tabela de parcelas
        StrSql = "SELECT * FROM tb_Crediario_Parcela where CodParc=" & VGIntCodParc
        RecParc.Open StrSql, vgCon, 1, 3
        
        RecParc("Quitado") = "sim"
        RecParc.Update
        
        'envia dados para o lan�amento de caixa
        StrSql = "SELECT * FROM tb_Caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        StrSql = "SELECT C.Nome FROM tb_Crediario as CR,tb_Cliente as C where CR.CodCli=C.CodCli and CR.CodCred=" & VGIntCodCred
        RecCli.Open StrSql, vgCon, 1, 3
    
        RecCxa.AddNew
        RecCxa("CodVenda") = 0
        RecCxa("DtMov") = FormataDataUS(TxtDtQuit.Text)
        RecCxa("TipoMov") = "Quita��o de parcela"
        RecCxa("Valor") = Mid(TxtParcPagar.Text, 4)
        RecCxa("TipoValor") = "credito"
        RecCxa("Descricao") = "Quita��o de parcela do credi�rio - Cliente: " & RecCli.Fields.Item(0).Value
        RecCxa("TipoPagto") = TipoPagto
        RecCxa.Update
        
        Desconecta
        
        VPStrBox = MsgBox("Parcela quitada.", vbInformation, "Pr� Vendas 2004 - Informa��o")
        
        'atualiza a grid de credi�rio
        FrmResumo_Crediario.MontaResumo
        
        Unload Me
        
        FrmResumo_Crediario.Enabled = True
        MDIPrincipal.Enabled = True
    End If
End Sub

Private Sub Form_Resize()
  FrmCrediario_Parcela_Quitar.Left = (MDIPrincipal.Width / 2) - (FrmCrediario_Parcela_Quitar.Width / 2)
  FrmCrediario_Parcela_Quitar.Top = (MDIPrincipal.Height / 3) - (FrmCrediario_Parcela_Quitar.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 4950
    Width = 5985
    
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    FrmResumo_Crediario.Enabled = False
    MDIPrincipal.Enabled = False
    
    If VGStrReceb = "parcela" Then
        VGStrReceb = ""
        VGIntCodParc = VGIntCodReceber
    End If
    
    Conecta
    
    Dim RecParc As New ADODB.Recordset
    
    StrSql = "SELECT Vencimento,Valor FROM tb_Crediario_Parcela WHERE CodParc=" & VGIntCodParc
    RecParc.Open StrSql, vgCon, 1, 3
    
    TxtValorParc.Text = FormataMoeda(VerificaNulo(RecParc!valor))
    TxtDtVenc.Text = FormataData(RecParc!vencimento)
    
    Desconecta
    
    '==== calcula se tem juros na parcela ====
    '==== 2% ao m�s e 0.033% ao dia ===========
    Dim QtdeMes As Double
    Dim QtdeDia As Double
    Dim JurosParc As Double
    Dim CalcJuros As Double
        
    If CDate(TxtDtVenc.Text) < Date Then
    'parcela vencida
        'calcula porcentagem de juros para a qtde de meses em atraso
        QtdeMes = DateDiff("m", TxtDtVenc.Text, Date) - 1
        
        'calcula porcentagem de juros para cada dia de atraso
        QtdeDia = DateDiff("d", TxtDtVenc.Text, Date)
        
        'calcula valor (em porcentagem) de juros que ser� cobrado
        CalcJuros = ArredondaNumDec((QtdeMes * 2) + (QtdeDia * 0.033))
        
        'calcula o valor da parcela acrescida do juros
        JurosParc = CCur((CCur(Mid(TxtValorParc.Text, 4)) * CalcJuros) / 100)
        
        TxtJuros.Text = CalcJuros
        TxtParcPagar.Text = FormataMoeda(CCur(Mid(TxtValorParc.Text, 4)) + JurosParc)
        
    Else
    'parcela em dia ou adiantada
        TxtParcPagar.Text = TxtValorParc.Text
    End If
    
    TxtDtQuit.Text = FormataData(Date)
    
End Sub

Private Sub OptPagtoChq_Click()
    FraPagtoChq.Visible = True
End Sub

Private Sub OptPagtoDin_Click()
    FraPagtoChq.Visible = False
End Sub

Private Sub TxtBanco_GotFocus()
    TxtBanco.SelStart = 0
    TxtBanco.SelLength = Len(TxtBanco.Text)
End Sub

Private Sub TxtCheque_GotFocus()
    TxtCheque.SelStart = 0
    TxtCheque.SelLength = Len(TxtCheque.Text)
End Sub

Private Sub TxtDesc_GotFocus()
    TxtDesc.SelStart = 0
    TxtDesc.SelLength = Len(TxtDesc.Text)
End Sub

Private Sub TxtDesc_LostFocus()
    Dim DescParc As Integer
    If TxtDesc.Text <> "" Then
        TxtJuros.Text = ""
        DescParc = CCur((CCur(TxtValorParc.Text) * TxtDesc.Text) / 100)
        TxtParcPagar.Text = FormataMoeda(CCur(TxtValorParc.Text) - DescParc)
        
    ElseIf TxtDesc.Text = "" And TxtJuros.Text = "" Then
        TxtParcPagar.Text = TxtValorParc.Text
    End If
End Sub

Private Sub TxtDigito_GotFocus()
    TxtDigito.SelStart = 0
    TxtDigito.SelLength = Len(TxtDigito.Text)
End Sub

Private Sub TxtDtQuit_GotFocus()
    If TxtDtQuit.Text = "__/__/____" Then
        TxtDtQuit.Text = ""
    End If
    TxtDtQuit.SelStart = 0
    TxtDtQuit.SelLength = Len(TxtDtQuit.Text)
End Sub

Private Sub TxtDtQuit_KeyPress(KeyAscii As Integer)
    '=== S� aceita n�meros ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtQuit_LostFocus()
    Dim VLStrData As String
    
    If TxtDtQuit.Text <> "" Then
        VLStrData = VerificaData(TxtDtQuit.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtQuit.SetFocus
        Else
            TxtDtQuit.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtQuit.Text = "__/__/____"
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
    '=== S� aceita n�meros ===
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

Private Sub TxtJuros_LostFocus()
    Dim JurosParc As Double
    
    If TxtJuros.Text <> "" Then
        TxtDesc.Text = ""
        JurosParc = CCur((CCur(Mid(TxtValorParc.Text, 4)) * TxtJuros.Text) / 100)
        TxtParcPagar.Text = FormataMoeda(CCur(Mid(TxtValorParc.Text, 4)) + JurosParc)
        
    ElseIf TxtJuros.Text = "" And TxtDesc.Text = "" Then
        TxtParcPagar.Text = TxtValorParc.Text
    End If
End Sub

Private Sub TxtParcPagar_GotFocus()
    TxtParcPagar.SelStart = 0
    TxtParcPagar.SelLength = Len(TxtParcPagar.Text)
End Sub

Private Sub TxtValorParc_GotFocus()
    TxtValorParc.SelStart = 0
    TxtValorParc.SelLength = Len(TxtValorParc.Text)
End Sub
