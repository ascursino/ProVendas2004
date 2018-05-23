VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmProduto_Lista_Imprimir 
   Caption         =   "Impressão do relatório de produtos"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
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
   Icon            =   "FrmProduto_Lista_Imprimir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4800
   Begin VB.Frame FraLista 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   4575
      Begin VB.OptionButton OptAtac 
         Caption         =   "relatório de produtos por preço de atacado"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         ToolTipText     =   "Relatório de produtos por preço de atacado"
         Top             =   960
         Width           =   4095
      End
      Begin VB.OptionButton OptVenda 
         Caption         =   "relatório de produtos por preço de venda"
         Height          =   255
         Left            =   360
         TabIndex        =   0
         ToolTipText     =   "Relatório de produtos por preço de venda"
         Top             =   600
         Width           =   3975
      End
      Begin VB.OptionButton OptCompl 
         Caption         =   "relatório de produtos completo"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         ToolTipText     =   "Relatório de produtos completo"
         Top             =   1320
         Width           =   3015
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmProduto_Lista_Imprimir.frx":0CCA
         TabIndex        =   7
         Top             =   240
         Width           =   3615
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
      TabIndex        =   5
      Top             =   1800
      Width           =   4575
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "&Imprimir"
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
         Left            =   1920
         TabIndex        =   3
         ToolTipText     =   "Imprimir relatório"
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   0
         OleObjectBlob   =   "FrmProduto_Lista_Imprimir.frx":0D6E
         Top             =   0
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
         Left            =   3240
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmProduto_Lista_Imprimir"
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

Private Sub CmdImprimir_Click()
    Screen.MousePointer = vbHourglass

    Dim nome As String
    Dim forn As String
    Dim tipo As String
    Dim desc As String
    Dim precofabric As String
    Dim precovendaunit As String
    Dim precovendaatac As String
    Dim moeda As String
    Dim VLStrLinha As String

    VLStrLinha = 1

    Conecta

    Do While VLStrLinha <= FrmPrincipal.GridProduto.MaxRows

        FrmPrincipal.GridProduto.Col = 1
        FrmPrincipal.GridProduto.Row = VLStrLinha
        nome = FrmPrincipal.GridProduto.Text

        FrmPrincipal.GridProduto.Col = 2
        FrmPrincipal.GridProduto.Row = VLStrLinha
        forn = FrmPrincipal.GridProduto.Text

        FrmPrincipal.GridProduto.Col = 3
        FrmPrincipal.GridProduto.Row = VLStrLinha
        tipo = FrmPrincipal.GridProduto.Text

        FrmPrincipal.GridProduto.Col = 4
        FrmPrincipal.GridProduto.Row = VLStrLinha
        desc = FrmPrincipal.GridProduto.Text

        FrmPrincipal.GridProduto.Col = 5
        FrmPrincipal.GridProduto.Row = VLStrLinha
        precofabric = FrmPrincipal.GridProduto.Text

        FrmPrincipal.GridProduto.Col = 6
        FrmPrincipal.GridProduto.Row = VLStrLinha
        precovendaunit = FrmPrincipal.GridProduto.Text

        FrmPrincipal.GridProduto.Col = 7
        FrmPrincipal.GridProduto.Row = VLStrLinha
        precovendaatac = FrmPrincipal.GridProduto.Text

        FrmPrincipal.GridProduto.Col = 8
        FrmPrincipal.GridProduto.Row = VLStrLinha
        moeda = FrmPrincipal.GridProduto.Text

        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08) " & _
        "VALUES ('" & nome & "','" & forn & "','" & tipo & "','" & desc & "','" & precofabric & "','" & precovendaunit & "','" & precovendaatac & "','" & moeda & "')"

        VLStrLinha = VLStrLinha + 1
    Loop

    Desconecta
    
    MDIPrincipal.Enabled = True
    
    If OptVenda.Value = True Then
        Unload Me
        rptProdutoVenda.Show
        
    ElseIf OptAtac.Value = True Then
        Unload Me
        rptProdutoAtac.Show
        
    ElseIf OptCompl.Value = True Then
        Unload Me
        rptProdutoCompl.Show
    End If

End Sub

Private Sub Form_Load()
    Height = 3135
    Width = 4920
    
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    OptCompl.Value = True
    'MDIPrincipal.Enabled = False

End Sub

Private Sub Form_Resize()
  FrmProduto_Lista_Imprimir.Left = (MDIPrincipal.Width / 2) - (FrmProduto_Lista_Imprimir.Width / 2)
  FrmProduto_Lista_Imprimir.Top = (MDIPrincipal.Height / 3) - (FrmProduto_Lista_Imprimir.Height / 3)
End Sub
