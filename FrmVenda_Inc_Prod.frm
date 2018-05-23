VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form FrmVenda_Inc_Prod 
   Caption         =   "Inclusão de Venda - Produto"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
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
   Icon            =   "FrmVenda_Inc_Prod.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   8640
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
      TabIndex        =   4
      Top             =   3960
      Width           =   8415
      Begin VB.CheckBox ChkPreco 
         Caption         =   "Usar preço atacado"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "Usar preço de atacado para este nesta venda"
         Top             =   360
         Width           =   2055
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   4560
         OleObjectBlob   =   "FrmVenda_Inc_Prod.frx":0CCA
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
         Left            =   7080
         TabIndex        =   2
         ToolTipText     =   "Fechar"
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
      Height          =   3855
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   8415
      Begin FPSpread.vaSpread GridProduto 
         Height          =   3495
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   8085
         _Version        =   393216
         _ExtentX        =   14261
         _ExtentY        =   6165
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
         MaxCols         =   5
         MaxRows         =   0
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "FrmVenda_Inc_Prod.frx":0EFE
         UserResize      =   1
      End
   End
End
Attribute VB_Name = "FrmVenda_Inc_Prod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdFechar_Click()
    VGIntCodProd = 0
    VGStrDescrProd = ""
    
    Unload Me
   
    FrmVenda_Inc.Enabled = True
End Sub

Private Sub Form_Resize()
  FrmVenda_Inc_Prod.Left = (MDIPrincipal.Width / 2) - (FrmVenda_Inc_Prod.Width / 2)
  FrmVenda_Inc_Prod.Top = (MDIPrincipal.Height / 2) - (FrmVenda_Inc_Prod.Height / 2.3)
End Sub

Private Sub Form_Load()
    Height = 5310
    Width = 8760
    
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    FrmVenda_Inc.Enabled = False
    
    Call MontaGridProduto
    
End Sub

Private Sub GridProduto_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim VLStrPreco As Currency
    
    GridProduto.Row = Row
    GridProduto.Col = 3
    If GridProduto.Text <> "" And GridProduto.Text <> "CodProd" Then
        VGIntCodProd = GridProduto.Text
        
        '=== Grid em venda ===
        FrmVenda_Inc.GridProduto.MaxRows = FrmVenda_Inc.GridProduto.MaxRows + 1
        FrmVenda_Inc.GridProduto.Row = FrmVenda_Inc.GridProduto.MaxRows
        
        '=== Grid em produto ===
        GridProduto.Row = GridProduto.ActiveRow
        
        'Produto
        FrmVenda_Inc.GridProduto.Col = 1
        GridProduto.Col = 1
        FrmVenda_Inc.GridProduto.Text = GridProduto.Text
        FrmVenda_Inc.GridProduto.Lock = True
        
        'Preço
        FrmVenda_Inc.GridProduto.Col = 2
        If ChkPreco.Value = 0 Then
            GridProduto.Col = 2
        Else
            GridProduto.Col = 3
        End If
        FrmVenda_Inc.GridProduto.Text = GridProduto.Text
        FrmVenda_Inc.GridProduto.Lock = True
        VLStrPreco = GridProduto.Text
        
        'CodProd
        FrmVenda_Inc.GridProduto.Col = 5
        GridProduto.Col = 5
        FrmVenda_Inc.GridProduto.Text = Val(GridProduto.Text)
        
        'Unload Me
        'FrmVenda_Inc.Enabled = True
        
    End If
End Sub

Sub MontaGridProduto()
    
    Dim VLIntLinha As Long
    Dim RecPesq As New ADODB.Recordset
    Dim RecForn As New ADODB.Recordset
    Dim RecEst As New ADODB.Recordset
    
    Conecta
    
    StrSql = "Select P.CodProd,P.CodForn,P.NomeProd,P.PrecoVendaUnit,P.PrecoVendaAtac " & _
             "from tb_Produto as P, tb_Estoque as E where P.CodProd=E.CodProd and E.QtdeProd <> 0"
    RecPesq.Open StrSql, vgCon, 1, 3
    
    If RecPesq.EOF Then
        GridProduto.Refresh
        GridProduto.MaxRows = 0
    Else
    
        VLIntLinha = 1
        GridProduto.MaxRows = VLIntLinha
         
        Do While Not RecPesq.EOF
                 
            GridProduto.Row = VLIntLinha
            GridProduto.Lock = True
            
            'Pegar fornecedor
            StrSql = "Select Nome from tb_Fornecedor where CodForn=" & RecPesq!CodForn
            RecForn.Open StrSql, vgCon, 1, 3
            
            'Produto
            GridProduto.Col = 1
            If Not RecForn.EOF Then
                GridProduto.Text = VerificaNulo(RecPesq!nomeprod) & " (" & VerificaNulo(RecForn!nome) & ")"
            Else
                GridProduto.Text = VerificaNulo(RecPesq!nomeprod) & " (?)"
            End If
            GridProduto.Lock = True
            
            'Preço Unit.
            GridProduto.Col = 2
            GridProduto.Text = FormataMoeda(VerificaNulo(RecPesq!precovendaunit))
            GridProduto.Lock = True
            
            'Preço Atac.
            GridProduto.Col = 3
            GridProduto.Text = FormataMoeda(VerificaNulo(RecPesq!precovendaatac))
            GridProduto.Lock = True
            
            'Qtde
            GridProduto.Col = 4
            StrSql = "Select QtdeProd from tb_Estoque where CodProd=" & RecPesq!CodProd
            RecEst.Open StrSql, vgCon, 1, 3

            If Not RecEst.EOF Then
                GridProduto.Text = FormataNum(RecEst!qtdeprod)
            Else
                GridProduto.Text = "0"
            End If
            GridProduto.Lock = True
            
            'CodProd
            GridProduto.Col = 5
            GridProduto.Text = Val(RecPesq!CodProd)
            GridProduto.Lock = True
            
            VLIntLinha = VLIntLinha + 1
            
            GridProduto.MaxRows = GridProduto.MaxRows + 1
            RecForn.Close
            RecEst.Close
            RecPesq.MoveNext
         Loop
         
         GridProduto.MaxRows = GridProduto.MaxRows - 1
    End If
    
    Desconecta
    
End Sub

