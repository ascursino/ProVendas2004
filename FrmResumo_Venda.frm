VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "spr32x60.ocx"
Begin VB.Form FrmResumo_Venda 
   Caption         =   "Resumo da Venda"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7905
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
   Icon            =   "FrmResumo_Venda.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   7905
   Begin VB.Frame FraProd 
      Caption         =   "Produtos"
      Height          =   2295
      Left            =   120
      TabIndex        =   46
      Top             =   480
      Width           =   7695
      Begin FPSpread.vaSpread GridListProd 
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   7365
         _Version        =   393216
         _ExtentX        =   12991
         _ExtentY        =   3201
         _StockProps     =   64
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   4
         MaxRows         =   0
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "FrmResumo_Venda.frx":0CCA
         UserResize      =   1
      End
   End
   Begin VB.Frame FraVista 
      Caption         =   "Finalização da venda / À vista"
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   7695
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel85 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Venda.frx":10EE
         TabIndex        =   4
         Top             =   480
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel86 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Venda.frx":1152
         TabIndex        =   5
         Top             =   840
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel87 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Venda.frx":11BC
         TabIndex        =   6
         Top             =   1200
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel88 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "FrmResumo_Venda.frx":1220
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblValorVista 
         Height          =   255
         Left            =   1200
         OleObjectBlob   =   "FrmResumo_Venda.frx":1292
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblDescVista 
         Height          =   255
         Left            =   1200
         OleObjectBlob   =   "FrmResumo_Venda.frx":1308
         TabIndex        =   9
         Top             =   840
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTotalVista 
         Height          =   255
         Left            =   1200
         OleObjectBlob   =   "FrmResumo_Venda.frx":1366
         TabIndex        =   10
         Top             =   1200
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTipoPagtoVista 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "FrmResumo_Venda.frx":13DE
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblBancoVista 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "FrmResumo_Venda.frx":1454
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblChequeVista 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "FrmResumo_Venda.frx":14C0
         TabIndex        =   13
         Top             =   960
         Width           =   2175
      End
   End
   Begin VB.Frame FraPrazoCheque 
      Caption         =   "Finalização da venda / A prazo - Cheque"
      Height          =   1815
      Left            =   120
      TabIndex        =   14
      Top             =   2880
      Visible         =   0   'False
      Width           =   7695
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel72 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Venda.frx":1540
         TabIndex        =   15
         Top             =   360
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel73 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Venda.frx":15A4
         TabIndex        =   16
         Top             =   600
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel74 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Venda.frx":160E
         TabIndex        =   17
         Top             =   840
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel75 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Venda.frx":1672
         TabIndex        =   18
         Top             =   1080
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblValorCheque 
         Height          =   255
         Left            =   1080
         OleObjectBlob   =   "FrmResumo_Venda.frx":16D6
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblParcCheque 
         Height          =   255
         Left            =   1080
         OleObjectBlob   =   "FrmResumo_Venda.frx":174C
         TabIndex        =   20
         Top             =   600
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblJurosCheque 
         Height          =   255
         Left            =   1080
         OleObjectBlob   =   "FrmResumo_Venda.frx":17B0
         TabIndex        =   21
         Top             =   840
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTotalCheque 
         Height          =   255
         Left            =   1080
         OleObjectBlob   =   "FrmResumo_Venda.frx":180E
         TabIndex        =   22
         Top             =   1080
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "FrmResumo_Venda.frx":1886
         TabIndex        =   25
         Top             =   360
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblValorParcCheque 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "FrmResumo_Venda.frx":18EE
         TabIndex        =   26
         Top             =   1440
         Width           =   2535
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblEntradaCheque 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "FrmResumo_Venda.frx":197A
         TabIndex        =   27
         Top             =   360
         Width           =   2775
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblBancoCheque 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "FrmResumo_Venda.frx":1A08
         TabIndex        =   28
         Top             =   600
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblValorEntradaCheque 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "FrmResumo_Venda.frx":1A74
         TabIndex        =   29
         Top             =   1200
         Width           =   2415
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblChequeCheque 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "FrmResumo_Venda.frx":1AFC
         TabIndex        =   30
         Top             =   840
         Width           =   2415
      End
   End
   Begin VB.Frame FraPrazoCarne 
      Caption         =   "Finalização da venda / A prazo - Carnê"
      Height          =   1815
      Left            =   120
      TabIndex        =   31
      Top             =   2880
      Visible         =   0   'False
      Width           =   7695
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Venda.frx":1B7C
         TabIndex        =   32
         Top             =   360
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Venda.frx":1BE0
         TabIndex        =   33
         Top             =   600
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Venda.frx":1C4A
         TabIndex        =   34
         Top             =   840
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Venda.frx":1CAE
         TabIndex        =   35
         Top             =   1080
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblValorCarne 
         Height          =   255
         Left            =   1080
         OleObjectBlob   =   "FrmResumo_Venda.frx":1D12
         TabIndex        =   36
         Top             =   360
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblParcCarne 
         Height          =   255
         Left            =   1080
         OleObjectBlob   =   "FrmResumo_Venda.frx":1D88
         TabIndex        =   37
         Top             =   600
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblJurosCarne 
         Height          =   255
         Left            =   1080
         OleObjectBlob   =   "FrmResumo_Venda.frx":1DEC
         TabIndex        =   38
         Top             =   840
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTotalCarne 
         Height          =   255
         Left            =   1080
         OleObjectBlob   =   "FrmResumo_Venda.frx":1E4A
         TabIndex        =   39
         Top             =   1080
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "FrmResumo_Venda.frx":1EC2
         TabIndex        =   40
         Top             =   360
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblValorParcCarne 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "FrmResumo_Venda.frx":1F2A
         TabIndex        =   41
         Top             =   1440
         Width           =   2535
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblEntradaCarne 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "FrmResumo_Venda.frx":1FB6
         TabIndex        =   42
         Top             =   360
         Width           =   2775
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblBancoCarne 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "FrmResumo_Venda.frx":2044
         TabIndex        =   43
         Top             =   600
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblValorEntradaCarne 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "FrmResumo_Venda.frx":20B0
         TabIndex        =   44
         Top             =   1200
         Width           =   2415
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblChequeCarne 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "FrmResumo_Venda.frx":2138
         TabIndex        =   45
         Top             =   840
         Width           =   2415
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
      TabIndex        =   2
      Top             =   4800
      Width           =   7695
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Venda.frx":21B8
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
         Left            =   6360
         TabIndex        =   0
         ToolTipText     =   "Fechar"
         Top             =   240
         Width           =   1095
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "FrmResumo_Venda.frx":23EC
      TabIndex        =   23
      Top             =   120
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblVendedor 
      Height          =   255
      Left            =   1200
      OleObjectBlob   =   "FrmResumo_Venda.frx":2456
      TabIndex        =   24
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "FrmResumo_Venda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrProd As String
Public VPStrQtde As String
Public VPStrVenda As String

Private Sub CmdFechar_Click()
    Unload Me
    
    MDIPrincipal.Enabled = True
    MDIPrincipal.WindowState = 2
    MDIPrincipal.SetFocus
End Sub

Private Sub Form_Resize()
  FrmResumo_Venda.Left = (MDIPrincipal.Width / 2) - (FrmResumo_Venda.Width / 2)
  FrmResumo_Venda.Top = (MDIPrincipal.Height / 3) - (FrmResumo_Venda.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 6135
    Width = 8025
    
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    MDIPrincipal.Enabled = False
        
    Conecta
    
    Dim RecVenda As New ADODB.Recordset
    Dim RecProd As New ADODB.Recordset
    Dim RecForn As New ADODB.Recordset
    Dim RecVendedor As New ADODB.Recordset
    Dim RecCred As New ADODB.Recordset
    Dim RecCredParc As New ADODB.Recordset
    
    StrSql = "Select * from tb_Venda_Produto where CodVenda=" & VGIntCodVenda
    RecVenda.Open StrSql, vgCon, 1, 3
    
    '==== Lista produto(s) =====
    VLIntLinha = 1
    GridListProd.MaxRows = VLIntLinha

    Do While Not RecVenda.EOF

        GridListProd.Row = VLIntLinha
        GridListProd.Lock = True

        'Produto
        GridListProd.Col = 1
        
        StrSql = "Select NomeProd from tb_Produto where CodProd=" & RecVenda!CodProd
        RecProd.Open StrSql, vgCon, 1, 3
        
        If Not RecProd.EOF Then
            GridListProd.Text = VerificaNulo(RecProd!nomeprod)
        Else
            GridListProd.Text = ""
        End If
        GridListProd.Lock = True

        'Fornecedor
        GridListProd.Col = 2
        
        StrSql = "Select F.Nome from tb_Fornecedor as F,tb_Produto as P where P.CodForn=F.CodForn and P.CodProd=" & RecVenda!CodProd
        RecForn.Open StrSql, vgCon, 1, 3
        
        If Not RecForn.EOF Then
            GridListProd.Text = VerificaNulo(RecForn!nome)
        Else
            GridListProd.Text = ""
        End If
        GridListProd.Lock = True

        'Qtde
        GridListProd.Col = 3
        GridListProd.Text = FormataNum(RecVenda!qtde)
        GridListProd.Lock = True

        'Valor venda
        GridListProd.Col = 4
        GridListProd.Text = FormataMoeda(RecVenda!valorprod)
        GridListProd.Lock = True

        VLIntLinha = VLIntLinha + 1

        GridListProd.MaxRows = GridListProd.MaxRows + 1
                    
        RecProd.Close
        RecForn.Close
        RecVenda.MoveNext
     Loop
     GridListProd.MaxRows = GridListProd.MaxRows - 1
     '==============================================================
     
    RecVenda.Close
     
    StrSql = "Select * from tb_Venda where CodVenda=" & VGIntCodVenda
    RecVenda.Open StrSql, vgCon, 1, 3
     
    StrSql = "Select Nome from tb_Vendedor where CodVendedor=" & RecVenda!CodVendedor
    RecVendedor.Open StrSql, vgCon, 1, 3
    
    If Not RecVendedor.EOF Then
        LblVendedor.Caption = RecVendedor!nome
    Else
        LblVendedor.Caption = ""
    End If
    
    StrSql = "Select ValorVenda,Parcela,Juros,ValorTotal,TipoEntr,ValorEntr,NumBanco,NumCheque from tb_Crediario where CodCred=" & RecVenda!CodCred
    RecCred.Open StrSql, vgCon, 1, 3
    
    StrSql = "Select Valor from tb_Crediario_Parcela where CodCred=" & RecVenda!CodCred
    RecCredParc.Open StrSql, vgCon, 1, 3
    
    If RecVenda!tipovenda = "A prazo - Cheque" Then
        LblValorCheque.Caption = FormataMoeda(RecCred!valorvenda)
        LblParcCheque.Caption = FormataNum(RecCred!parcela)
        
        If RecCred!juros <> "" And IsNull(RecCred!juros) = False Then
            LblJurosCheque.Caption = FormataNum(RecCred!juros) & "%"
        Else
            LblJurosCheque.Caption = ""
        End If
        
        LblTotalCheque.Caption = FormataMoeda(RecCred!valortotal)
        LblEntradaCheque.Caption = VerificaNulo(RecCred!tipoentr)
        
        If RecCred!tipoentr = "Cheque" Then
            LblBancoCheque.Caption = "Banco: " & VerificaNulo(RecCred!numbanco)
            LblChequeCheque.Caption = "Cheque: " & VerificaNulo(RecCred!numcheque)
            LblValorEntradaCheque.Caption = "Valor entrada: " & FormataMoeda(RecCred!valorentr)
            LblBancoCheque.Visible = True
            LblChequeCheque.Visible = True
            LblValorEntradaCheque.Visible = True
            
        ElseIf RecCred!tipoentr = "Dinheiro" Then
            LblValorEntradaCheque.Caption = "Valor entrada: " & FormataMoeda(RecCred!valorentr)
            LblBancoCheque.Visible = False
            LblChequeCheque.Visible = False
            LblValorEntradaCheque.Visible = True
        Else
            LblBancoCheque.Visible = False
            LblChequeCheque.Visible = False
            LblValorEntradaCheque.Visible = False
        End If
        
        LblValorParcCheque.Caption = FormataNum(RecCred!parcela) & " parcela(s) de " & FormataMoeda(RecCredParc!valor)
        
        FraPrazoCheque.Visible = True
        FraPrazoCarne.Visible = False
        FraVista.Visible = False
        
    ElseIf RecVenda!tipovenda = "A prazo - Carnê" Then
        LblValorCarne.Caption = FormataMoeda(RecCred!valorvenda)
        LblParcCarne.Caption = FormataNum(RecCred!parcela)
        
        If RecCred!juros <> "" And IsNull(RecCred!juros) = False Then
            LblJurosCarne.Caption = FormataNum(RecCred!juros) & "%"
        Else
            LblJurosCarne.Caption = ""
        End If
        
        LblJurosCarne.Caption = FormataNum(RecCred!juros) & "%"
        LblTotalCarne.Caption = FormataMoeda(RecCred!valortotal)
        LblEntradaCarne.Caption = VerificaNulo(RecCred!tipoentr)

        If RecCred!tipoentr = "Cheque" Then
            LblBancoCarne.Caption = "Banco: " & VerificaNulo(RecCred!numbanco)
            LblChequeCarne.Caption = "Cheque: " & VerificaNulo(RecCred!numcheque)
            LblValorEntradaCarne.Caption = "Valor entrada: " & FormataMoeda(RecCred!valorentr)
            LblBancoCarne.Visible = True
            LblChequeCarne.Visible = True
            LblValorEntradaCarne.Visible = True

        ElseIf RecCred!tipoentr = "Dinheiro" Then
            LblValorEntradaCarne.Caption = "Valor entrada: " & FormataMoeda(RecCred!valorentr)
            LblBancoCarne.Visible = False
            LblChequeCarne.Visible = False
            LblValorEntradaCarne.Visible = True
        Else
            LblBancoCarne.Visible = False
            LblChequeCarne.Visible = False
            LblValorEntradaCarne.Visible = False
        End If

        LblValorParcCarne.Caption = FormataNum(RecCred!parcela) & " parcela(s) de " & FormataMoeda(RecCredParc!valor)

        FraPrazoCheque.Visible = False
        FraPrazoCarne.Visible = True
        FraVista.Visible = False

    ElseIf RecVenda!tipovenda = "À vista" Then
        LblValorVista.Caption = FormataMoeda(RecVenda!subtotalvenda)
        
        If RecVenda!desconto <> "" And IsNull(RecVenda!desconto) = False Then
            LblDescVista.Caption = FormataNum(RecVenda!desconto) & "%"
        Else
            LblDescVista.Caption = ""
        End If
        
        LblTotalVista.Caption = FormataMoeda(RecVenda!totalvenda)
        LblTipoPagtoVista.Caption = VerificaNulo(RecVenda!TipoPagto)
        
        If RecVenda!TipoPagto = "Dinheiro" Then
            LblBancoVista.Visible = False
            LblChequeVista.Visible = False
        Else
            LblBancoVista.Caption = "Banco: " & VerificaNulo(RecVenda!numbanco)
            LblChequeVista.Caption = "Cheque: " & VerificaNulo(RecVenda!numcheque)
            LblBancoVista.Visible = True
            LblChequeVista.Visible = True
        End If
        
        FraPrazoCheque.Visible = False
        FraPrazoCarne.Visible = False
        FraVista.Visible = True
    
    End If
   
    Desconecta
End Sub

