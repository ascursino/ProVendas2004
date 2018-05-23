VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "spr32x60.ocx"
Begin VB.Form FrmResumo_Orcamento 
   Caption         =   "Resumo do Orçamento"
   ClientHeight    =   7800
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
   Icon            =   "FrmResumo_Orcamento.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
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
      Height          =   6855
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   7095
      Begin VB.TextBox TxtObs 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   22
         Text            =   "FrmResumo_Orcamento.frx":0CCA
         Top             =   3960
         Width           =   6615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmResumo_Orcamento.frx":0CD4
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmResumo_Orcamento.frx":0D3C
         TabIndex        =   5
         Top             =   1320
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmResumo_Orcamento.frx":0DAC
         TabIndex        =   6
         Top             =   1680
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   3840
         OleObjectBlob   =   "FrmResumo_Orcamento.frx":0E22
         TabIndex        =   7
         Top             =   2400
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   4200
         OleObjectBlob   =   "FrmResumo_Orcamento.frx":0E8A
         TabIndex        =   8
         Top             =   2760
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Left            =   3840
         OleObjectBlob   =   "FrmResumo_Orcamento.frx":0EFA
         TabIndex        =   9
         Top             =   2040
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmResumo_Orcamento.frx":0F6E
         TabIndex        =   10
         Top             =   3360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmResumo_Orcamento.frx":0FDC
         TabIndex        =   11
         Top             =   3720
         Width           =   1215
      End
      Begin FPSpread.vaSpread GridListProd 
         Height          =   1815
         Left            =   240
         TabIndex        =   1
         Top             =   4800
         Width           =   6615
         _Version        =   393216
         _ExtentX        =   11668
         _ExtentY        =   3201
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
         MaxCols         =   3
         MaxRows         =   0
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   2
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "FrmResumo_Orcamento.frx":104A
         UserResize      =   1
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmResumo_Orcamento.frx":13D5
         TabIndex        =   12
         Top             =   2040
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTotalVista 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "FrmResumo_Orcamento.frx":1449
         TabIndex        =   13
         Top             =   2040
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTotalVenda 
         Height          =   255
         Left            =   1800
         OleObjectBlob   =   "FrmResumo_Orcamento.frx":14A9
         TabIndex        =   14
         Top             =   1680
         Width           =   2295
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTel 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Orcamento.frx":1509
         TabIndex        =   15
         Top             =   1320
         Width           =   3735
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli 
         Height          =   255
         Left            =   1080
         OleObjectBlob   =   "FrmResumo_Orcamento.frx":1569
         TabIndex        =   16
         Top             =   600
         Width           =   5775
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTotalPrazo 
         Height          =   255
         Left            =   5160
         OleObjectBlob   =   "FrmResumo_Orcamento.frx":15C9
         TabIndex        =   17
         Top             =   2040
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblQtdeParc 
         Height          =   255
         Left            =   3840
         OleObjectBlob   =   "FrmResumo_Orcamento.frx":1629
         TabIndex        =   18
         Top             =   2760
         Width           =   255
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblValorParc 
         Height          =   255
         Left            =   5400
         OleObjectBlob   =   "FrmResumo_Orcamento.frx":1685
         TabIndex        =   19
         Top             =   2760
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblEntrada 
         Height          =   255
         Left            =   4800
         OleObjectBlob   =   "FrmResumo_Orcamento.frx":16E7
         TabIndex        =   20
         Top             =   2400
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblValidade 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "FrmResumo_Orcamento.frx":1749
         TabIndex        =   21
         Top             =   3360
         Width           =   5535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmResumo_Orcamento.frx":17AB
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblData 
         Height          =   255
         Left            =   1080
         OleObjectBlob   =   "FrmResumo_Orcamento.frx":180D
         TabIndex        =   24
         Top             =   240
         Width           =   5775
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmResumo_Orcamento.frx":186D
         TabIndex        =   25
         Top             =   960
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblVendedor 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "FrmResumo_Orcamento.frx":18D7
         TabIndex        =   26
         Top             =   960
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
      TabIndex        =   2
      Top             =   6960
      Width           =   7095
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
         TabIndex        =   0
         ToolTipText     =   "Fechar"
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   360
         OleObjectBlob   =   "FrmResumo_Orcamento.frx":1939
         Top             =   120
      End
   End
End
Attribute VB_Name = "FrmResumo_Orcamento"
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
  FrmResumo_Orcamento.Left = (MDIPrincipal.Width / 2) - (FrmResumo_Orcamento.Width / 2)
  FrmResumo_Orcamento.Top = (MDIPrincipal.Height / 3) - (FrmResumo_Orcamento.Height / 3)
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
        
    Height = 8310
    Width = 7440
    
    Call MontaResumo
    
    MDIPrincipal.Enabled = False
End Sub

Sub MontaResumo()
        
    FrmPrincipal.GridOrcamento.Row = FrmPrincipal.GridOrcamento.ActiveRow
    
    'Data
    FrmPrincipal.GridOrcamento.Col = 1
    LblData.Caption = FrmPrincipal.GridOrcamento.Text
    
    'Cliente
    FrmPrincipal.GridOrcamento.Col = 2
    LblCli.Caption = FrmPrincipal.GridOrcamento.Text
    
    'Vendedor
    FrmPrincipal.GridOrcamento.Col = 3
    LblVendedor.Caption = FrmPrincipal.GridOrcamento.Text
    
    'Telefone
    FrmPrincipal.GridOrcamento.Col = 4
    LblTel.Caption = FrmPrincipal.GridOrcamento.Text
    
    'Total da venda e Total à vista
    FrmPrincipal.GridOrcamento.Col = 5
    LblTotalVenda.Caption = FrmPrincipal.GridOrcamento.Text
    LblTotalVista.Caption = FrmPrincipal.GridOrcamento.Text
    
    'Total a prazo
    FrmPrincipal.GridOrcamento.Col = 9
    LblTotalPrazo.Caption = FrmPrincipal.GridOrcamento.Text

    'Entrada
    FrmPrincipal.GridOrcamento.Col = 7
    LblEntrada.Caption = FrmPrincipal.GridOrcamento.Text

    'Qtde Parcela
    FrmPrincipal.GridOrcamento.Col = 6
    LblQtdeParc.Caption = FrmPrincipal.GridOrcamento.Text

    'Valor Parccela
    FrmPrincipal.GridOrcamento.Col = 8
    LblValorParc.Caption = FrmPrincipal.GridOrcamento.Text

    'Validade
    FrmPrincipal.GridOrcamento.Col = 10
    LblValidade.Caption = FrmPrincipal.GridOrcamento.Text

    'Observação
    FrmPrincipal.GridOrcamento.Col = 11
    TxtObs.Text = FrmPrincipal.GridOrcamento.Text

    'Produtos
    Dim VLIntLinha As Integer
    Dim RecOrcProd As New ADODB.Recordset
    Dim RecProd As New ADODB.Recordset
    
    VLIntLinha = 1
    GridListProd.MaxRows = VLIntLinha
    
    Conecta
    
    StrSql = "Select * from tb_Orcamento_Produto where CodOrc=" & VGIntCodOrc
    RecOrcProd.Open StrSql, vgCon, 1, 3
    
    Do While Not RecOrcProd.EOF
        StrSql = "Select NomeProd from tb_Produto where CodProd=" & RecOrcProd!CodProd
        RecProd.Open StrSql, vgCon, 1, 3

        GridListProd.Row = VLIntLinha
        GridListProd.Lock = True

        'Produto
        GridListProd.Col = 1
        GridListProd.Text = RecProd!nomeprod
        GridListProd.Lock = True

        'Valor
        GridListProd.Col = 2
        GridListProd.Text = FormataMoeda(RecOrcProd!valorprod)
        GridListProd.Lock = True

        'Qtde
        GridListProd.Col = 3
        GridListProd.Text = FormataNum(RecOrcProd!qtde)
        GridListProd.Lock = True
    
        VLIntLinha = VLIntLinha + 1

        GridListProd.MaxRows = GridListProd.MaxRows + 1
        
        RecProd.Close
        
        RecOrcProd.MoveNext
     Loop
    
     Desconecta
    
     GridListProd.MaxRows = GridListProd.MaxRows - 1
End Sub
