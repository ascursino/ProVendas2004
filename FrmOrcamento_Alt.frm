VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form FrmOrcamento_Alt 
   Caption         =   "Alteração de Orçamento"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7320
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
   Icon            =   "FrmOrcamento_Alt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   7320
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
      Height          =   7095
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   7095
      Begin VB.TextBox TxtTel2 
         Height          =   285
         Left            =   4920
         MaxLength       =   15
         TabIndex        =   2
         ToolTipText     =   "Número do telefone do cliente"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox TxtTotalPrazo 
         Height          =   285
         Left            =   5280
         MaxLength       =   10
         TabIndex        =   5
         ToolTipText     =   "Total da venda a prazo"
         Top             =   3480
         Width           =   1575
      End
      Begin VB.TextBox TxtValorParc 
         Height          =   285
         Left            =   5280
         MaxLength       =   10
         TabIndex        =   8
         ToolTipText     =   "Valor das parcelas"
         Top             =   4920
         Width           =   1575
      End
      Begin VB.TextBox TxtEntrada 
         Height          =   285
         Left            =   5280
         MaxLength       =   10
         TabIndex        =   7
         ToolTipText     =   "Valor da entrada"
         Top             =   4440
         Width           =   1575
      End
      Begin VB.TextBox TxtTotalVista 
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   4
         ToolTipText     =   "Total da venda à vista"
         Top             =   3480
         Width           =   1575
      End
      Begin VB.ComboBox CboVendedor 
         Height          =   315
         ItemData        =   "FrmOrcamento_Alt.frx":0CCA
         Left            =   1320
         List            =   "FrmOrcamento_Alt.frx":0CCC
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Nome do vendedor"
         Top             =   5520
         Width           =   5535
      End
      Begin VB.ComboBox CboQtdeParc 
         Height          =   315
         ItemData        =   "FrmOrcamento_Alt.frx":0CCE
         Left            =   5280
         List            =   "FrmOrcamento_Alt.frx":0CD0
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Quantidade de parcelas"
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox TxtObs 
         Height          =   645
         Left            =   240
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "Observação sobre o cliente e/ou orçamento"
         Top             =   6240
         Width           =   6615
      End
      Begin VB.TextBox TxtTel1 
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   1
         ToolTipText     =   "Número do telefone do cliente"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox TxtCli 
         Height          =   285
         Left            =   1320
         MaxLength       =   200
         TabIndex        =   0
         ToolTipText     =   "Nome do cliente"
         Top             =   240
         Width           =   5535
      End
      Begin VB.TextBox TxtValidade 
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   9
         Text            =   "__/__/____"
         ToolTipText     =   "Data da validade"
         Top             =   5040
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmOrcamento_Alt.frx":0CD2
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmOrcamento_Alt.frx":0D3A
         TabIndex        =   17
         Top             =   720
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmOrcamento_Alt.frx":0DA8
         TabIndex        =   18
         Top             =   3480
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "FrmOrcamento_Alt.frx":0E1C
         TabIndex        =   19
         Top             =   3960
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "FrmOrcamento_Alt.frx":0E96
         TabIndex        =   20
         Top             =   4440
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   3840
         OleObjectBlob   =   "FrmOrcamento_Alt.frx":0EFE
         TabIndex        =   21
         Top             =   4920
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "FrmOrcamento_Alt.frx":0F6E
         TabIndex        =   22
         Top             =   3480
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmOrcamento_Alt.frx":0FE2
         TabIndex        =   23
         Top             =   5040
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmOrcamento_Alt.frx":1050
         TabIndex        =   24
         Top             =   5520
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmOrcamento_Alt.frx":10BA
         TabIndex        =   25
         Top             =   6000
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   3840
         OleObjectBlob   =   "FrmOrcamento_Alt.frx":1128
         TabIndex        =   26
         Top             =   720
         Width           =   975
      End
      Begin FPSpread.vaSpread GridListProd 
         Height          =   2055
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   6615
         _Version        =   393216
         _ExtentX        =   11668
         _ExtentY        =   3625
         _StockProps     =   64
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   5
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "FrmOrcamento_Alt.frx":1196
         UserResize      =   1
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblQtdeParc 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "FrmOrcamento_Alt.frx":163B
         TabIndex        =   27
         Top             =   4920
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
      TabIndex        =   14
      Top             =   7200
      Width           =   7095
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   120
         OleObjectBlob   =   "FrmOrcamento_Alt.frx":1697
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
         Left            =   5760
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
         Left            =   4440
         TabIndex        =   12
         ToolTipText     =   "Efetuar alteração"
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmOrcamento_Alt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String
Public VPIntEntrada As Currency

Private Sub CboQtdeParc_Click()
    If CboQtdeParc.Text <> "" And CboQtdeParc.Text <> "00" Then
        TxtEntrada.Text = FormataMoeda(CCur(TxtTotalPrazo.Text) / Int(CboQtdeParc.Text))
        TxtValorParc.Text = FormataMoeda(CCur(TxtTotalPrazo.Text) / Int(CboQtdeParc.Text))
        LblQtdeParc.Caption = FormataNum(CboQtdeParc.Text - 1)
    End If
End Sub

Private Sub CmdFechar_Click()
    Unload Me
    
    MDIPrincipal.Enabled = True
    MDIPrincipal.WindowState = 2
End Sub

Private Sub CmdOK_Click()
    
    Conecta
    
    Dim RecOrc As New ADODB.Recordset
    Dim RecVend As New ADODB.Recordset
    Dim RecProd As New ADODB.Recordset
    Dim RecVerif As New ADODB.Recordset
    
    Dim VLIntCodOrc As Integer
    Dim VLIntRows As Integer
    Dim VLCurTotalProd As Currency
    Dim VLIntQtde As Integer
    Dim VLCurValor As Currency
    Dim VLIntCodProd As Integer
    Dim VLStrVend As String
    
    'altera na tabela de orçamento
    StrSql = "SELECT * FROM tb_Orcamento where CodOrc=" & VGIntCodOrc
    RecOrc.Open StrSql, vgCon, 1, 3
        
    RecOrc("CodVendedor") = Trim(Mid(CboVendedor.Text, Len(CboVendedor.Text) - 10))
    RecOrc("Nome") = TxtCli.Text
    RecOrc("Telefone1") = TxtTel1.Text
    RecOrc("Telefone2") = TxtTel2.Text
    RecOrc("TotalVenda") = CCur(TxtTotalVista.Text)
    If CboQtdeParc.Text = "" And CboQtdeParc.Text = "00" Then
        RecOrc("Parcela") = 0
    Else
        RecOrc("Parcela") = CboQtdeParc.Text
    End If
    RecOrc("Entrada") = CCur(TxtEntrada.Text)
    RecOrc("ValorParc") = CCur(TxtValorParc.Text)
    RecOrc("ValorPrazo") = CCur(TxtTotalPrazo.Text)
    RecOrc("Validade") = FormataDataUS(TxtValidade.Text)
    RecOrc("Obs") = Trim(TxtObs.Text)
    RecOrc.Update
        
    RecOrc.Close
    
    'altera na tabela de orçamento_produto
    StrSql = "SELECT * FROM tb_Orcamento_Produto where CodOrc=" & VGIntCodOrc
    RecOrc.Open StrSql, vgCon, 1, 3
    
    VLIntRows = GridListProd.MaxRows
    
    Do While VLIntRows <> 0
        GridListProd.Row = VLIntRows
        GridListProd.Col = 1
        
        If GridListProd.Text = "1" Then
            GridListProd.Row = VLIntRows
            GridListProd.Col = 5
            VLIntCodProd = GridListProd.Text
            
            GridListProd.Row = VLIntRows
            GridListProd.Col = 3
            VLCurValor = GridListProd.Text
            
            GridListProd.Col = 4
            VLIntQtde = GridListProd.Text
            
            VLCurTotalProd = VLCurValor * VLIntQtde
            
            StrSql = "SELECT * FROM tb_Orcamento_Produto where CodOrc=" & VGIntCodOrc & " and CodProd=" & VLIntCodProd
            RecVerif.Open StrSql, vgCon, 1, 3
            
            If Not RecVerif.EOF Then
                'altera produto já existente no orçamento
                RecOrc("ValorProd") = VLCurValor
                RecOrc("Qtde") = VLIntQtde
                RecOrc("ValorTotalProd") = VLCurTotalProd
                RecOrc.Update
            Else
                'insere novo produto no orçamento
                vgCon.Execute ("Insert into tb_Orcamento_Produto (CodOrc,CodProd,ValorProd,Qtde,ValorTotalProd) " & _
                              "values(" & VGIntCodOrc & "," & VLIntCodProd & ",'" & VLCurValor & "'," & VLIntQtde & ",'" & VLCurTotalProd & "')")
                
            End If
            RecVerif.Close
            
        Else
            GridListProd.Row = VLIntRows
            GridListProd.Col = 5
            VLIntCodProd = GridListProd.Text
        
            StrSql = "SELECT * FROM tb_Orcamento_Produto where CodOrc=" & VGIntCodOrc & " and CodProd=" & VLIntCodProd
            RecVerif.Open StrSql, vgCon, 1, 3
            
            If Not RecVerif.EOF Then
                vgCon.Execute ("Delete from tb_Orcamento_Produto where CodOrc=" & VGIntCodOrc & " and CodProd=" & VLIntCodProd)
            End If
            RecVerif.Close
        End If
        
        VLIntRows = VLIntRows - 1
    Loop
    
    Desconecta
    
    VPStrBox = MsgBox("Alteração efetuada.", vbInformation, "Pró Vendas 2004 - Informação")
    
    FrmPrincipal.CmdPesqOrc.Value = True
    
    Unload Me
    
    MDIPrincipal.Enabled = True
    MDIPrincipal.WindowState = 2
        
End Sub

Private Sub Form_Resize()
  FrmOrcamento_Alt.Left = (MDIPrincipal.Width / 2) - (FrmOrcamento_Alt.Width / 2)
  FrmOrcamento_Alt.Top = (MDIPrincipal.Height / 3) - (FrmOrcamento_Alt.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 8535
    Width = 7440
    
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    MDIPrincipal.Enabled = False
    
    Call MontaCbos
    
    Conecta
    
    Dim RecOrc As New ADODB.Recordset
    Dim RecVend As New ADODB.Recordset
    
    StrSql = "SELECT * FROM tb_Orcamento where CodOrc=" & VGIntCodOrc
    RecOrc.Open StrSql, vgCon, 1, 3
    
    TxtCli.Text = VerificaNulo(RecOrc!nome)
    TxtTel1.Text = VerificaNulo(RecOrc!telefone1)
    TxtTel2.Text = VerificaNulo(RecOrc!telefone2)
    TxtTotalVista.Text = FormataMoeda(VerificaNulo(RecOrc!totalvenda))
    TxtTotalPrazo.Text = FormataMoeda(VerificaNulo(RecOrc!valorprazo))
    CboQtdeParc.Text = FormataNum(RecOrc!parcela)
    TxtEntrada.Text = FormataMoeda(VerificaNulo(RecOrc!entrada))
    TxtValorParc.Text = FormataMoeda(VerificaNulo(RecOrc!valorparc))
    TxtValidade.Text = FormataData(RecOrc!validade)
    TxtObs.Text = VerificaNulo(RecOrc!obs)
    
    StrSql = "SELECT * FROM tb_Vendedor where CodVendedor=" & RecOrc!CodVendedor
    RecVend.Open StrSql, vgCon, 1, 3
    
    If Not RecVend.EOF Then
        CboVendedor.Text = RecVend!nome & "                                                                                                      " & RecVend!CodVendedor
    End If
    
    Desconecta
    
    Call MontaGrid

End Sub

Sub MontaCbos()
    Conecta
    
    Dim RecVend As New ADODB.Recordset
    Dim count As Integer
    
    'monta cbo de vendedor
    StrSql = "SELECT CodVendedor,Nome FROM tb_Vendedor order by Nome"
    RecVend.Open StrSql, vgCon, 1, 3
    
    CboVendedor.AddItem ("                                                                                                                 0")
    Do While Not RecVend.EOF
        CboVendedor.AddItem (RecVend!nome & "                                                                                                      " & RecVend!CodVendedor)
        RecVend.MoveNext
    Loop
    
    Desconecta
    
    'monta cbo de quantidade de parcela
    count = 0
    Do While count <= 24
        CboQtdeParc.AddItem (FormataNum(count))
        count = count + 1
    Loop
    
End Sub

Sub MontaGrid()
    Dim VLIntLinha As Integer
    Dim RecOrcProd As New ADODB.Recordset
    Dim RecProd As New ADODB.Recordset
    
    VLIntLinha = 1
    GridListProd.MaxRows = VLIntLinha
    
    Conecta
    
    StrSql = "Select * from tb_Produto"
    RecProd.Open StrSql, vgCon, 1, 3
    
    Do While Not RecProd.EOF

        StrSql = "Select * from tb_Orcamento_Produto where CodOrc=" & VGIntCodOrc & " and CodProd=" & RecProd!CodProd
        RecOrcProd.Open StrSql, vgCon, 1, 3

        If Not RecOrcProd.EOF Then
            'produto existe neste orçamento
            GridListProd.Row = VLIntLinha
            GridListProd.Lock = True
    
            'Checkbox
            GridListProd.Col = 1
            GridListProd.CellType = CellTypeCheckBox
            GridListProd.TypeCheckType = TypeCheckTypeNormal
            GridListProd.TypeCheckCenter = True
            GridListProd.Text = "1"
            GridListProd.Lock = False
            
            'Produto
            GridListProd.Col = 2
            GridListProd.Text = RecProd!nomeprod
            GridListProd.Lock = True
    
            'Valor
            GridListProd.Col = 3
            GridListProd.Text = FormataMoeda(RecOrcProd!valorprod)
            GridListProd.Lock = True
    
            'Qtde
            GridListProd.Col = 4
            GridListProd.Text = FormataNum(RecOrcProd!qtde)
            GridListProd.Lock = False
    
            'CodProd
            GridListProd.Col = 5
            GridListProd.Text = Val(RecOrcProd!CodProd)
            GridListProd.Lock = True
        Else
            'produto NÃO existe neste orçamento
            GridListProd.Row = VLIntLinha
            GridListProd.Lock = True
    
            'Checkbox
            GridListProd.Col = 1
            GridListProd.CellType = CellTypeCheckBox
            GridListProd.TypeCheckType = TypeCheckTypeNormal
            GridListProd.TypeCheckCenter = True
            GridListProd.Lock = False
            
            'Produto
            GridListProd.Col = 2
            GridListProd.Text = RecProd!nomeprod
            GridListProd.Lock = True
    
            'Valor
            GridListProd.Col = 3
            GridListProd.Text = FormataMoeda(RecProd!precovendaunit)
            GridListProd.Lock = True
    
            'Qtde
            GridListProd.Col = 4
            GridListProd.Text = ""
            GridListProd.Lock = False
    
            'CodProd
            GridListProd.Col = 5
            GridListProd.Text = Val(RecProd!CodProd)
            GridListProd.Lock = True
        End If
        
        VLIntLinha = VLIntLinha + 1

        GridListProd.MaxRows = GridListProd.MaxRows + 1
        
        RecOrcProd.Close
        
        RecProd.MoveNext
     Loop
    
     Desconecta
    
     GridListProd.MaxRows = GridListProd.MaxRows - 1
End Sub

Private Sub GridListProd_LostFocus()
    Dim VLIntRows As Integer
    Dim VLCurTotal As Currency
    Dim VLIntQtde As Integer
    Dim VLCurValor As Currency
    
    VLIntRows = GridListProd.MaxRows
    
    Do While VLIntRows <> 0
        GridListProd.Row = VLIntRows
        GridListProd.Col = 1
        
        If GridListProd.Text = "1" Then
            GridListProd.Row = VLIntRows
            GridListProd.Col = 3
            VLCurValor = GridListProd.Text
            
            GridListProd.Col = 4
            If GridListProd.Text = "" Then
                VLIntQtde = 0
            Else
                VLIntQtde = GridListProd.Text
            End If
            
            VLCurTotal = VLCurTotal + (VLCurValor * VLIntQtde)
        End If
        
        VLIntRows = VLIntRows - 1
    Loop
    
    TxtTotalVista.Text = FormataMoeda(VLCurTotal)
    TxtTotalPrazo.Text = FormataMoeda(VLCurTotal)
    CboQtdeParc.SetFocus
    Call CboQtdeParc_Click
End Sub

Private Sub TxtCli_GotFocus()
    TxtCli.SelStart = 0
    TxtCli.SelLength = Len(TxtCli.Text)
End Sub

Private Sub TxtEntrada_GotFocus()
    TxtEntrada.SelStart = 0
    TxtEntrada.SelLength = Len(TxtEntrada.Text)
End Sub

Private Sub TxtEntrada_LostFocus()
    If TxtEntrada.Text <> "" Then
        If TxtTotalPrazo.Text <> "" And CboQtdeParc.Text <> "" And CboQtdeParc.Text <> "00" Then
            TxtEntrada.Text = FormataMoeda(TxtEntrada.Text)
            VLIntRestante = CCur(TxtTotalPrazo.Text) - CCur(TxtEntrada.Text)
            TxtValorParc.Text = FormataMoeda(CCur(VLIntRestante) / Int(CboQtdeParc.Text - 1))
            LblQtdeParc.Caption = FormataNum(CboQtdeParc.Text - 1)
        End If
    Else
        If TxtTotalPrazo.Text <> "" And CboQtdeParc.Text <> "" Then
            TxtValorParc.Text = FormataMoeda(CCur(TxtTotalPrazo.Text) / Int(CboQtdeParc.Text))
        End If
        
        If TxtTotalPrazo.Text <> "" And CboQtdeParc.Text <> "" And CboQtdeParc.Text <> "00" Then
            TxtEntrada.Text = ""
            TxtValorParc.Text = FormataMoeda(CCur(TxtTotalPrazo.Text) / Int(CboQtdeParc.Text))
            LblQtdeParc.Caption = FormataNum(CboQtdeParc.Text)
        End If
    End If
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
    '=== Só aceita números, parênteses, espaço e - ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtTel2_GotFocus()
    TxtTel2.SelStart = 0
    TxtTel2.SelLength = Len(TxtTel2.Text)
End Sub

Private Sub TxtTel2_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, parênteses, espaço e - ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtTotalPrazo_GotFocus()
    TxtTotalPrazo.SelStart = 0
    TxtTotalPrazo.SelLength = Len(TxtTotalPrazo.Text)
End Sub

Private Sub TxtTotalVista_GotFocus()
    TxtTotalVista.SelStart = 0
    TxtTotalVista.SelLength = Len(TxtTotalVista.Text)
End Sub

Private Sub TxtTotalVista_LostFocus()
    If TxtTotalVista.Text <> "" Then
        TxtTotalVista.Text = FormataMoeda(TxtTotalVista.Text)
    End If
End Sub

Private Sub TxtTotalPrazo_LostFocus()
    If TxtTotalPrazo.Text <> "" Then
        TxtTotalPrazo.Text = FormataMoeda(TxtTotalPrazo.Text)
    End If
End Sub

Private Sub TxtValorParc_GotFocus()
    TxtValorParc.SelStart = 0
    TxtValorParc.SelLength = Len(TxtValorParc.Text)
End Sub

Private Sub TxtValorParc_LostFocus()
    If TxtValorParc.Text <> "" And TxtValorParc.Text <> "R$" Then
        If TxtTotalPrazo.Text <> "" And CboQtdeParc.Text <> "" And CboQtdeParc.Text <> "00" Then
            TxtEntrada.Text = FormataMoeda(CCur(TxtTotalPrazo.Text) - (CCur(TxtValorParc.Text) * Int(LblQtdeParc.Caption)))
            TxtValorParc.Text = FormataMoeda(TxtValorParc.Text)
        End If
    End If
End Sub

Private Sub TxtValidade_GotFocus()
    If TxtValidade.Text = "__/__/____" Then
        TxtValidade.Text = ""
    End If
    TxtValidade.SelStart = 0
    TxtValidade.SelLength = Len(TxtValidade.Text)
End Sub

Private Sub TxtValidade_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e barra ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtValidade_LostFocus()
    Dim VLStrData As String
    
    If TxtValidade.Text <> "" Then
        VLStrData = VerificaData(TxtValidade.Text)
        
        If VGStrDataErro = "sim" Then
            TxtValidade.SetFocus
        Else
            TxtValidade.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtValidade.Text = "__/__/____"
    End If
End Sub
