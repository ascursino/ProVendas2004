VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "spr32x60.ocx"
Begin VB.Form FrmResumo_Crediario 
   Caption         =   "Resumo do Crediário"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6510
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
   Icon            =   "FrmResumo_Crediario.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   6510
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
      Height          =   5775
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   6255
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmResumo_Crediario.frx":0CCA
         TabIndex        =   7
         Top             =   600
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmResumo_Crediario.frx":0D32
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmResumo_Crediario.frx":0DA8
         TabIndex        =   9
         Top             =   1680
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmResumo_Crediario.frx":0E18
         TabIndex        =   10
         Top             =   3480
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmResumo_Crediario.frx":0E8E
         TabIndex        =   11
         Top             =   3120
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmResumo_Crediario.frx":0F02
         TabIndex        =   12
         Top             =   960
         Width           =   1095
      End
      Begin FPSpread.vaSpread GridParcela 
         Height          =   1815
         Left            =   120
         TabIndex        =   4
         Top             =   3840
         Width           =   6015
         _Version        =   393216
         _ExtentX        =   10610
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
         SpreadDesigner  =   "FrmResumo_Crediario.frx":0F72
         UserResize      =   1
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmResumo_Crediario.frx":13EC
         TabIndex        =   13
         Top             =   2760
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTipoEntr 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "FrmResumo_Crediario.frx":145E
         TabIndex        =   14
         Top             =   2760
         Width           =   4575
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblValorVenda 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "FrmResumo_Crediario.frx":14BE
         TabIndex        =   15
         Top             =   1680
         Width           =   4575
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTipoCred 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "FrmResumo_Crediario.frx":151E
         TabIndex        =   16
         Top             =   1320
         Width           =   4575
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli 
         Height          =   255
         Left            =   1080
         OleObjectBlob   =   "FrmResumo_Crediario.frx":157E
         TabIndex        =   17
         Top             =   600
         Width           =   5055
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblValorEntr 
         Height          =   255
         Left            =   1680
         OleObjectBlob   =   "FrmResumo_Crediario.frx":15DE
         TabIndex        =   18
         Top             =   3120
         Width           =   4455
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumParc 
         Height          =   255
         Left            =   1680
         OleObjectBlob   =   "FrmResumo_Crediario.frx":163E
         TabIndex        =   19
         Top             =   3480
         Width           =   4455
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCredsta 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Crediario.frx":16A0
         TabIndex        =   20
         Top             =   960
         Width           =   4695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmResumo_Crediario.frx":1702
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblData 
         Height          =   255
         Left            =   1080
         OleObjectBlob   =   "FrmResumo_Crediario.frx":1764
         TabIndex        =   22
         Top             =   240
         Width           =   5055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmResumo_Crediario.frx":17C4
         TabIndex        =   23
         Top             =   2040
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmResumo_Crediario.frx":1828
         TabIndex        =   24
         Top             =   2400
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblValorTotal 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmResumo_Crediario.frx":1898
         TabIndex        =   25
         Top             =   2400
         Width           =   4695
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblJuros 
         Height          =   255
         Left            =   960
         OleObjectBlob   =   "FrmResumo_Crediario.frx":18F8
         TabIndex        =   26
         Top             =   2040
         Width           =   5175
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
      Top             =   5880
      Width           =   6255
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
         Left            =   840
         TabIndex        =   0
         ToolTipText     =   "Alterar parcela"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdExcluir 
         Caption         =   "&Excluir"
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
         Left            =   2160
         TabIndex        =   1
         ToolTipText     =   "Excluir parcela"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdQuitar 
         Caption         =   "&Quitar"
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
         TabIndex        =   2
         ToolTipText     =   "Quitar parcela"
         Top             =   240
         Width           =   1215
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
         TabIndex        =   3
         ToolTipText     =   "Fechar"
         Top             =   240
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   120
         OleObjectBlob   =   "FrmResumo_Crediario.frx":1958
         Top             =   120
      End
   End
End
Attribute VB_Name = "FrmResumo_Crediario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAlterar_Click()
    FrmCrediario_Parcela_Alt.Show
End Sub

Private Sub CmdExcluir_Click()
    VPStrResponse = MsgBox("Deseja excluir esta parcela?", vbYesNo, "Pró Vendas 2004 - Informação")

    If VPStrResponse = vbYes Then
        Conecta
        vgCon.Execute ("DELETE FROM tb_Crediario_Parcela WHERE CodParc=" & VGIntCodParc)
        Desconecta
        
        Call MontaResumo
    End If
End Sub

Private Sub CmdFechar_Click()
    Unload Me
    
    MDIPrincipal.Enabled = True
    MDIPrincipal.WindowState = 2
    MDIPrincipal.SetFocus
End Sub

Private Sub CmdQuitar_Click()
    FrmCrediario_Parcela_Quitar.Show
End Sub

Private Sub Form_Resize()
  FrmResumo_Crediario.Left = (MDIPrincipal.Width / 2) - (FrmResumo_Crediario.Width / 2)
  FrmResumo_Crediario.Top = (MDIPrincipal.Height / 3) - (FrmResumo_Crediario.Height / 3)
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
        
    Height = 7215
    Width = 6630
    
    Call MontaResumo
    
    MDIPrincipal.Enabled = False
End Sub

Sub MontaResumo()
        
    FrmPrincipal.GridCrediario.Row = FrmPrincipal.GridCrediario.ActiveRow
    
    'Data
    FrmPrincipal.GridCrediario.Col = 3
    LblData.Caption = FrmPrincipal.GridCrediario.Text
    
    'Cliente
    FrmPrincipal.GridCrediario.Col = 1
    LblCli.Caption = FrmPrincipal.GridCrediario.Text
    
    'Crediarista
    FrmPrincipal.GridCrediario.Col = 2
    LblCredsta.Caption = FrmPrincipal.GridCrediario.Text
    
    'Tipo crediário
    FrmPrincipal.GridCrediario.Col = 4
    LblTipoCred.Caption = FrmPrincipal.GridCrediario.Text
    
    'Valor venda
    FrmPrincipal.GridCrediario.Col = 5
    LblValorVenda.Caption = FrmPrincipal.GridCrediario.Text
    
    'Juros
    FrmPrincipal.GridCrediario.Col = 6
    LblJuros.Caption = FrmPrincipal.GridCrediario.Text

    'Valor total
    FrmPrincipal.GridCrediario.Col = 7
    LblValorTotal.Caption = FrmPrincipal.GridCrediario.Text

    'Valor Entrada
    FrmPrincipal.GridCrediario.Col = 9
    LblValorEntr.Caption = FrmPrincipal.GridCrediario.Text

    'Tipo entrada
    FrmPrincipal.GridCrediario.Col = 8
    LblTipoEntr.Caption = FrmPrincipal.GridCrediario.Text

    'Nº de Parcelas
    FrmPrincipal.GridCrediario.Col = 10
    LblNumParc.Caption = FrmPrincipal.GridCrediario.Text

    'Parcelas
    Dim VLIntLinha As Integer
    Dim RecParc As New ADODB.Recordset
    
    VLIntLinha = 1
    GridParcela.MaxRows = VLIntLinha
    
    Conecta
    
    StrSql = "Select CodParc,NumParc,Vencimento,Valor,Quitado from tb_Crediario_Parcela where CodCred=" & VGIntCodCred
    RecParc.Open StrSql, vgCon, 1, 3
    
    Do While Not RecParc.EOF
        GridParcela.Row = VLIntLinha
        GridParcela.Lock = True
        
        'Parcela
        GridParcela.Col = 1
        GridParcela.Text = FormataNum(RecParc!numparc) & "/" & LblNumParc.Caption
        GridParcela.Lock = True
        
        'Vencimento
        GridParcela.Col = 2
        GridParcela.Text = FormataData(RecParc!vencimento)
        GridParcela.Lock = True
        
        'Valor
        GridParcela.Col = 3
        GridParcela.Text = FormataMoeda(RecParc!valor)
        GridParcela.Lock = True
        
        'Quitado
        GridParcela.Col = 4
        GridParcela.Text = RecParc!quitado
        GridParcela.Lock = True
        
        'CodParc
        GridParcela.Col = 5
        GridParcela.Text = Val(RecParc!codparc)
        GridParcela.Lock = True
        
        VLIntLinha = VLIntLinha + 1

        GridParcela.MaxRows = GridParcela.MaxRows + 1
        
        RecParc.MoveNext
     Loop
    
     Desconecta
    
     GridParcela.MaxRows = GridParcela.MaxRows - 1
End Sub

Private Sub GridParcela_Click(ByVal Col As Long, ByVal Row As Long)
    GridParcela.Row = Row
    GridParcela.Col = 5
    If GridParcela.Text <> "" And GridParcela.Text <> "CodParc" Then
        VGIntCodParc = GridParcela.Text
        CmdAlterar.Enabled = True
        CmdExcluir.Enabled = True
        
        GridParcela.Row = Row
        GridParcela.Col = 4
        If GridParcela.Text = "sim" Then
            CmdQuitar.Enabled = False
        Else
            CmdQuitar.Enabled = True
        End If
        
    Else
        CmdAlterar.Enabled = False
        CmdExcluir.Enabled = False
        CmdQuitar.Enabled = False
    End If
End Sub

Private Sub GridParcela_DblClick(ByVal Col As Long, ByVal Row As Long)
    GridParcela.Row = Row
    GridParcela.Col = 5
    If GridParcela.Text <> "" And GridParcela.Text <> "CodParc" Then
        VGIntCodParc = GridParcela.Text
        FrmResumo_Parcela.Show
    End If
End Sub
