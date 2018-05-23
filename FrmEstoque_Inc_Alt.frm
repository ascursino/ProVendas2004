VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form FrmEstoque_Inc_Alt 
   Caption         =   "Inclusão/Alteração de Produtos no Estoque"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
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
   Icon            =   "FrmEstoque_Inc_Alt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   8760
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
      Top             =   4440
      Width           =   8535
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   2280
         OleObjectBlob   =   "FrmEstoque_Inc_Alt.frx":0CCA
         Top             =   240
      End
      Begin VB.CommandButton CmdAlterar 
         Caption         =   "&Alterar"
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
         Left            =   5880
         TabIndex        =   7
         ToolTipText     =   "Alterar produto no estoque"
         Top             =   240
         Width           =   1095
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
         Left            =   7200
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CmdIncluir 
         Caption         =   "&Incluir"
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
         Left            =   4560
         TabIndex        =   6
         ToolTipText     =   "Incluir produto no estoque"
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
      Height          =   4215
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   8535
      Begin VB.TextBox TxtQtdeProd 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   5640
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "(QtdeProd)"
         ToolTipText     =   "Descrição do produto"
         Top             =   2880
         Width           =   2775
      End
      Begin VB.TextBox TxtUltPed 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   5640
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "(UltPed)"
         ToolTipText     =   "Última quantidade inserida no estoque"
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox TxtQtdeMin 
         Height          =   285
         Left            =   7440
         MaxLength       =   4
         TabIndex        =   4
         ToolTipText     =   "Quantide mínima recomendada do produto no estoque"
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox TxtProd 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   765
         Left            =   5640
         MultiLine       =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Text            =   "FrmEstoque_Inc_Alt.frx":0EFE
         ToolTipText     =   "Descrição do produto"
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox TxtQtdeEst 
         Height          =   285
         Left            =   7440
         MaxLength       =   4
         TabIndex        =   5
         ToolTipText     =   "Quantidade do produto para estoque"
         Top             =   3720
         Width           =   975
      End
      Begin FPSpread.vaSpread GridProduto 
         Height          =   3855
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5415
         _Version        =   393216
         _ExtentX        =   9551
         _ExtentY        =   6800
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
         MaxCols         =   2
         MaxRows         =   0
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "FrmEstoque_Inc_Alt.frx":0F0A
         UserResize      =   1
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   5640
         OleObjectBlob   =   "FrmEstoque_Inc_Alt.frx":125C
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   5640
         OleObjectBlob   =   "FrmEstoque_Inc_Alt.frx":12C4
         TabIndex        =   12
         Top             =   3360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   5640
         OleObjectBlob   =   "FrmEstoque_Inc_Alt.frx":1334
         TabIndex        =   13
         Top             =   3720
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   495
         Left            =   5640
         OleObjectBlob   =   "FrmEstoque_Inc_Alt.frx":13B0
         TabIndex        =   14
         Top             =   1440
         Width           =   2775
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   495
         Left            =   5640
         OleObjectBlob   =   "FrmEstoque_Inc_Alt.frx":1454
         TabIndex        =   15
         ToolTipText     =   "Quantidade do produto em estoque"
         Top             =   2400
         Width           =   2655
      End
   End
End
Attribute VB_Name = "FrmEstoque_Inc_Alt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String

Private Sub CmdAlterar_Click()
    Dim RecProd As New ADODB.Recordset
    
    If TxtQtdeMin.Text = "" Or TxtQtdeEst.Text = "" Then
        VPStrBox = MsgBox("Preencha todos os campos.", vbInformation, "Pró Vendas 2004 - Informação")
    Else
        Conecta
        
        StrSql = "Select * from tb_Estoque where CodProd=" & VGIntCodProd
        RecProd.Open StrSql, vgCon, 1, 3
        
        If TxtQtdeMin.Text = "" Or TxtQtdeMin.Text = "0" Then
            RecProd("QtdeMin") = 0
        Else
            RecProd("QtdeMin") = TxtQtdeMin.Text
        End If
        
        If TxtQtdeEst.Text <> "" And TxtQtdeEst.Text <> "0" Then
            RecProd("QtdeProd") = RecProd!qtdeprod + TxtQtdeEst.Text
            RecProd("UltPed") = TxtQtdeEst.Text
        End If
        RecProd.Update
        
        Desconecta
        
        VPStrBox = MsgBox("Alteração efetuada.", vbInformation, "Pró Vendas 2004 - Informação")
        
        TxtProd.Text = ""
        TxtUltPed.Text = ""
        TxtQtdeProd.Text = ""
        TxtQtdeMin.Text = ""
        TxtQtdeEst.Text = ""
    End If
End Sub

Private Sub CmdFechar_Click()
    Unload Me
    
    MDIPrincipal.Enabled = True
    MDIPrincipal.WindowState = 2
End Sub

Private Sub CmdIncluir_Click()
    Dim RecProd As New ADODB.Recordset
    
    If TxtQtdeMin.Text = "" Or TxtQtdeEst.Text = "" Then
        VPStrBox = MsgBox("Preencha todos os campos.", vbInformation, "Pró Vendas 2004 - Informação")
    Else
        Conecta
        
        StrSql = "Select * from tb_Estoque"
        RecProd.Open StrSql, vgCon, 1, 3
        
        RecProd.AddNew
        RecProd("CodProd") = VGIntCodProd
        
        If TxtQtdeMin.Text = "" Then
            RecProd("QtdeMin") = 0
        Else
            RecProd("QtdeMin") = TxtQtdeMin.Text
        End If
        
        If TxtQtdeEst.Text = "" Then
            RecProd("QtdeProd") = 0
            RecProd("UltPed") = 0
        Else
            RecProd("QtdeProd") = TxtQtdeEst.Text
            RecProd("UltPed") = TxtQtdeEst.Text
        End If
        RecProd.Update
        
        Desconecta
        
        VPStrBox = MsgBox("Estoque cadastrado.", vbInformation, "Pró Vendas 2004 - Informação")
        
        TxtProd.Text = ""
        TxtUltPed.Text = ""
        TxtQtdeProd.Text = ""
        TxtQtdeMin.Text = ""
        TxtQtdeEst.Text = ""
    End If
End Sub

Private Sub Form_Resize()
  FrmEstoque_Inc_Alt.Left = (MDIPrincipal.Width / 2) - (FrmEstoque_Inc_Alt.Width / 2)
  FrmEstoque_Inc_Alt.Top = (MDIPrincipal.Height / 3) - (FrmEstoque_Inc_Alt.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 5790
    Width = 8880
    
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    MDIPrincipal.Enabled = False
    
    CmdIncluir.Enabled = False
    CmdAlterar.Enabled = False
    
    TxtProd.Text = ""
    TxtUltPed.Text = ""
    TxtQtdeProd.Text = ""
    
    Call MontaGridProduto
End Sub

Sub MontaGridProduto()
    Dim VLIntCodProd As Long
    Dim VLIntLinha As Long
    Dim RecProd As New ADODB.Recordset
     
    Conecta
    
    StrSql = "Select CodProd,NomeProd from tb_Produto order by NomeProd"
    RecProd.Open StrSql, vgCon, 1, 3
    
    If RecProd.EOF Then
        CmdIncluir.Enabled = False
        CmdAlterar.Enabled = False
    Else
        VLIntLinha = 1
        GridProduto.MaxRows = VLIntLinha
         
        Do While Not RecProd.EOF
            GridProduto.Row = VLIntLinha
            GridProduto.Lock = True
            
            'Produto
            GridProduto.Col = 1
            GridProduto.Text = VerificaNulo(RecProd!nomeprod)
            GridProduto.Lock = True
            
            'CodProd
            GridProduto.Col = 2
            GridProduto.Text = Val(RecProd!CodProd)
            GridProduto.Lock = True
            
            VLIntLinha = VLIntLinha + 1
            
            GridProduto.MaxRows = GridProduto.MaxRows + 1
            RecProd.MoveNext
         Loop
         
         GridProduto.MaxRows = GridProduto.MaxRows - 1
    End If
    
    Desconecta
    
End Sub

Private Sub GridProduto_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim RecProd As New ADODB.Recordset
    Dim RecPreco As New ADODB.Recordset
    Dim VLStrProduto As String
    
    GridProduto.Row = Row
    GridProduto.Col = 2
    
    If GridProduto.Text <> "" And GridProduto.Text <> "CodProd" Then
        VGIntCodProd = GridProduto.Text
        
        GridProduto.Row = Row
        GridProduto.Col = 1
        VLStrProduto = GridProduto.Text
        
        Conecta
        
        StrSql = "Select QtdeMin,QtdeProd,UltPed from tb_Estoque where CodProd=" & VGIntCodProd
        RecProd.Open StrSql, vgCon, 1, 3
        
        If Not RecProd.EOF Then
            TxtProd.Text = VLStrProduto
            TxtUltPed.Text = FormataNum(RecProd!ultped)
            TxtQtdeProd.Text = FormataNum(RecProd!qtdeprod)
            TxtQtdeMin.Text = FormataNum(RecProd!qtdemin)
            
            TxtQtdeEst.SetFocus
            
            CmdIncluir.Enabled = False
            CmdAlterar.Enabled = True
        Else
            TxtProd.Text = VLStrProduto
            TxtUltPed.Text = ""
            TxtQtdeProd.Text = ""
            TxtQtdeMin.Text = ""
            
            VPStrBox = MsgBox("Não existe informação de estoque cadastrado para este produto." & Chr(13) & "Se preferir, pode cadastrar agora.", vbInformation, "Pró Vendas 2004 - Informação")
            
            TxtQtdeMin.SetFocus
            
            CmdIncluir.Enabled = True
            CmdAlterar.Enabled = False
        End If
    Else
        CmdIncluir.Enabled = False
        CmdAlterar.Enabled = False
    End If
    
    Desconecta
    
End Sub

Private Sub TxtQtdeEst_GotFocus()
    TxtQtdeEst.SelStart = 0
    TxtQtdeEst.SelLength = Len(TxtQtdeEst.Text)
End Sub

Private Sub TxtQtdeEst_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtQtdeEst_LostFocus()
    If TxtQtdeEst.Text = "" Or TxtQtdeEst.Text = "0" Then
        TxtQtdeEst.Text = 0
    Else
        TxtQtdeEst.Text = FormataNum(TxtQtdeEst.Text)
    End If
End Sub

Private Sub TxtQtdeMin_GotFocus()
    TxtQtdeMin.SelStart = 0
    TxtQtdeMin.SelLength = Len(TxtQtdeMin.Text)
End Sub

Private Sub TxtQtdeMin_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtQtdeMin_LostFocus()
    If TxtQtdeMin.Text = "" Or TxtQtdeMin.Text = "0" Then
        TxtQtdeMin.Text = 0
    Else
        TxtQtdeMin.Text = FormataNum(TxtQtdeMin.Text)
    End If
End Sub
