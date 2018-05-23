VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmAssinaturaOrc 
   Caption         =   "Personalização de orçamento"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
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
   Icon            =   "FrmAssinaturaOrc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   6480
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
      TabIndex        =   11
      Top             =   3000
      Width           =   6255
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
         Left            =   4680
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   240
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   120
         OleObjectBlob   =   "FrmAssinaturaOrc.frx":0CCA
         Top             =   120
      End
      Begin VB.CommandButton CmdBranco 
         Caption         =   "&Em branco"
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
         Left            =   3120
         TabIndex        =   8
         ToolTipText     =   "Personalizar orçamento em branco"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdOK 
         Caption         =   "&Personalizar"
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
         Left            =   1440
         TabIndex        =   7
         ToolTipText     =   "Personalizar orçamento"
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   6255
      Begin VB.TextBox TxtLogo 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         ToolTipText     =   "Logomarca da empresa"
         Top             =   600
         Width           =   4935
      End
      Begin VB.CommandButton CmdListaArq 
         Caption         =   "&Listar arquivos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   1
         ToolTipText     =   "Listar arquivos de sua máquina"
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox TxtWeb 
         Height          =   285
         Left            =   1200
         MaxLength       =   60
         TabIndex        =   6
         ToolTipText     =   "Endereço do site o email da empresa"
         Top             =   2400
         Width           =   4935
      End
      Begin VB.TextBox TxtNome 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   2
         ToolTipText     =   "Nome da empresa"
         Top             =   960
         Width           =   4935
      End
      Begin VB.TextBox TxtTel 
         Height          =   285
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   5
         ToolTipText     =   "Telefone da empresa"
         Top             =   2040
         Width           =   4935
      End
      Begin VB.TextBox TxtBairro 
         Height          =   285
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   4
         ToolTipText     =   "Bairro da empresa"
         Top             =   1680
         Width           =   4935
      End
      Begin VB.TextBox TxtEndereco 
         Height          =   285
         Left            =   1200
         MaxLength       =   60
         TabIndex        =   3
         ToolTipText     =   "Endereço da empresa"
         Top             =   1320
         Width           =   4935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmAssinaturaOrc.frx":0EFE
         TabIndex        =   12
         Top             =   1320
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmAssinaturaOrc.frx":0F68
         TabIndex        =   13
         Top             =   1680
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmAssinaturaOrc.frx":0FCE
         TabIndex        =   14
         Top             =   2040
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmAssinaturaOrc.frx":1038
         TabIndex        =   15
         Top             =   960
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmAssinaturaOrc.frx":10A0
         TabIndex        =   16
         Top             =   2400
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmAssinaturaOrc.frx":1100
         TabIndex        =   17
         Top             =   600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmAssinaturaOrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String

Private Sub CmdBranco_Click()
    Unload Me
    If VGStrAssinaturaProposta = "proposta" Then
        VGStrAssinaturaProp = "branco"
    Else
        VGStrAssinaturaOrc = "branco"
    End If
    
    MDIPrincipal.Enabled = True
    
    If VGStrAssinaturaProposta = "proposta" Then
        VGStrAssinaturaProposta = ""
        FrmVenda_Inc.MontaImpressaoProposta
    Else
        Call ImprimirOrc
    End If
    VGStrPersonalizar = ""
End Sub

Private Sub CmdFechar_Click()
    VGStrAssinaturaProposta = ""
    Unload Me
    MDIPrincipal.Enabled = True
    VGStrPersonalizar = ""
End Sub

Private Sub CmdListaArq_Click()
    FrmListaArquivos.Show
End Sub

Private Sub CmdOK_Click()
    If TxtLogo.Text = "" And TxtNome.Text = "" And TxtEndereco.Text = "" And TxtBairro.Text = "" And TxtTel.Text = "" And TxtWeb.Text = "" Then
        If VGStrAssinaturaProposta = "proposta" Then
            VPStrBox = MsgBox("Preencha os campos para personalizar a proposta de crédito." & Chr(13) & "Caso não deseje personalizar escolha o botão 'Em Branco'", vbInformation, "Pró Vendas 2004 - Informação")
        Else
            VPStrBox = MsgBox("Preencha os campos para personalizar o orçamento." & Chr(13) & "Caso não deseje personalizar escolha o botão 'Em Branco'", vbInformation, "Pró Vendas 2004 - Informação")
        End If
    Else
        Conecta
        
        Dim RecAss As New ADODB.Recordset
        
        If VGStrAssinaturaProposta = "proposta" Then
            StrSql = "Select * From tb_AssinaturaProp"
        Else
            StrSql = "Select * From tb_AssinaturaOrc"
        End If
        
        RecAss.Open StrSql, vgCon, 1, 3
        
        If RecAss.EOF Then
            RecAss.AddNew
        End If
        
        If VGStrAssinaturaProposta <> "proposta" Then
            RecAss("Logo") = TxtLogo.Text
        End If
        RecAss("Nome") = TxtNome.Text
        RecAss("Endereco") = TxtEndereco.Text
        RecAss("Bairro") = TxtBairro.Text
        RecAss("Telefone") = TxtTel.Text
        RecAss("Web") = TxtWeb.Text
        RecAss.Update
        
        Desconecta
        
        MDIPrincipal.Enabled = True
        
        Unload Me
        
        If VGStrAssinaturaProposta = "proposta" Then
            VGStrAssinaturaProp = "personalizada"
            VGStrAssinaturaProposta = ""
            FrmVenda_Inc.MontaImpressaoProposta
        Else
            VGStrAssinaturaOrc = "personalizada"
            Call ImprimirOrc
        End If
    End If
    VGStrPersonalizar = ""
End Sub

Private Sub Form_Resize()
  FrmAssinaturaOrc.Left = (MDIPrincipal.Width / 2) - (FrmAssinaturaOrc.Width / 2)
  FrmAssinaturaOrc.Top = (MDIPrincipal.Height / 3) - (FrmAssinaturaOrc.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 4350
    Width = 6600
    
    MDIPrincipal.Enabled = False
    
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    Conecta
    
    Dim RecAss As New ADODB.Recordset
    
    If VGStrAssinaturaProposta = "proposta" Then
        Me.Caption = "Personalização da proposta de crédito"
        TxtLogo.Enabled = False
        CmdListaArq.Enabled = False
        StrSql = "Select * From tb_AssinaturaProp"
    Else
        Me.Caption = "Personalização de orçamento"
        TxtLogo.Enabled = True
        CmdListaArq.Enabled = True
        StrSql = "Select * From tb_AssinaturaOrc"
    End If
    
    RecAss.Open StrSql, vgCon, 1, 3
    
    If Not RecAss.EOF Then
        If VGStrAssinaturaProposta <> "proposta" Then
            If IsNull(RecAss!Logo) = False Then
                TxtLogo.Text = RecAss!Logo
            End If
        End If
        
        If IsNull(RecAss!nome) = False Then
            TxtNome.Text = RecAss!nome
        End If
        
        If IsNull(RecAss!endereco) = False Then
            TxtEndereco.Text = RecAss!endereco
        End If
        
        If IsNull(RecAss!bairro) = False Then
            TxtBairro.Text = RecAss!bairro
        End If
        
        If IsNull(RecAss!telefone) = False Then
            TxtTel.Text = RecAss!telefone
        End If
        
        If IsNull(RecAss!web) = False Then
            TxtWeb.Text = RecAss!web
        End If
    End If
    
    Desconecta
End Sub

Private Sub TxtBairro_GotFocus()
    TxtBairro.SelStart = 0
    TxtBairro.SelLength = Len(TxtBairro.Text)
End Sub

Private Sub TxtEndereco_GotFocus()
    TxtEndereco.SelStart = 0
    TxtEndereco.SelLength = Len(TxtEndereco.Text)
End Sub

Private Sub TxtNome_GotFocus()
    TxtNome.SelStart = 0
    TxtNome.SelLength = Len(TxtNome.Text)
End Sub

Private Sub TxtLogo_GotFocus()
    TxtLogo.SelStart = 0
    TxtLogo.SelLength = Len(TxtLogo.Text)
End Sub

Private Sub TxtTel_GotFocus()
    TxtTel.SelStart = 0
    TxtTel.SelLength = Len(TxtTel.Text)
End Sub

Private Sub TxtTel_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, parênteses, espaço e - ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtWeb_GotFocus()
    TxtWeb.SelStart = 0
    TxtWeb.SelLength = Len(TxtWeb.Text)
End Sub

Sub ImprimirOrc()
    Screen.MousePointer = vbHourglass

    Dim data As String
    Dim cliente As String
    Dim vendedor As String
    Dim telefone As String
    Dim totalvista As String
    Dim parcelado As String
    Dim entrada As String
    Dim valorparc As String
    Dim totalprazo As String
    Dim validade As String
    Dim obs As String
    Dim CodOrc As Integer

    Dim VLStrLinha As String
    Dim RecProd As New ADODB.Recordset
    Dim RecOrcProd As New ADODB.Recordset
    
    VLStrLinha = FrmPrincipal.GridOrcamento.ActiveRow
    
    Conecta

    FrmPrincipal.GridOrcamento.Col = 1
    FrmPrincipal.GridOrcamento.Row = VLStrLinha
    data = FrmPrincipal.GridOrcamento.Text

    FrmPrincipal.GridOrcamento.Col = 2
    FrmPrincipal.GridOrcamento.Row = VLStrLinha
    cliente = FrmPrincipal.GridOrcamento.Text

    FrmPrincipal.GridOrcamento.Col = 3
    FrmPrincipal.GridOrcamento.Row = VLStrLinha
    vendedor = FrmPrincipal.GridOrcamento.Text

    FrmPrincipal.GridOrcamento.Col = 4
    FrmPrincipal.GridOrcamento.Row = VLStrLinha
    telefone = FrmPrincipal.GridOrcamento.Text

    FrmPrincipal.GridOrcamento.Col = 5
    FrmPrincipal.GridOrcamento.Row = VLStrLinha
    totalvista = FrmPrincipal.GridOrcamento.Text

    FrmPrincipal.GridOrcamento.Col = 6
    FrmPrincipal.GridOrcamento.Row = VLStrLinha
    parcelado = FormataNum(Val(Trim(Mid(FrmPrincipal.GridOrcamento.Text, 1, InStr(FrmPrincipal.GridOrcamento.Text, " ")))) - 1)

    FrmPrincipal.GridOrcamento.Col = 7
    FrmPrincipal.GridOrcamento.Row = VLStrLinha
    entrada = FrmPrincipal.GridOrcamento.Text

    FrmPrincipal.GridOrcamento.Col = 8
    FrmPrincipal.GridOrcamento.Row = VLStrLinha
    valorparc = FrmPrincipal.GridOrcamento.Text

    FrmPrincipal.GridOrcamento.Col = 9
    FrmPrincipal.GridOrcamento.Row = VLStrLinha
    totalprazo = FrmPrincipal.GridOrcamento.Text

    FrmPrincipal.GridOrcamento.Col = 10
    FrmPrincipal.GridOrcamento.Row = VLStrLinha
    validade = FrmPrincipal.GridOrcamento.Text

    FrmPrincipal.GridOrcamento.Col = 11
    FrmPrincipal.GridOrcamento.Row = VLStrLinha
    obs = FrmPrincipal.GridOrcamento.Text

    FrmPrincipal.GridOrcamento.Col = 12
    FrmPrincipal.GridOrcamento.Row = VLStrLinha
    CodOrc = FrmPrincipal.GridOrcamento.Text

    StrSql = "Select * From tb_Orcamento_Produto where CodOrc=" & CodOrc
    RecOrcProd.Open StrSql, vgCon, 1, 3

    Do While Not RecOrcProd.EOF
        StrSql = "Select NomeProd From tb_Produto where CodProd=" & RecOrcProd!CodProd
        RecProd.Open StrSql, vgCon, 1, 3
    
        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08,campo09,campo10,campo11,campo12,campo13,campo14,campo15,campo16) " & _
        "VALUES ('" & data & "','" & cliente & "','" & telefone & "',''," & _
        "'" & vendedor & "','" & totalvista & "','" & totalprazo & "'," & _
        "'" & entrada & "','" & parcelado & "','" & valorparc & "'," & _
        "'" & validade & "','" & obs & "','" & RecProd!nomeprod & "'," & _
        "'" & FormataMoeda(RecOrcProd!valorprod) & "','" & RecOrcProd!qtde & "'," & _
        "'" & FormataMoeda(RecOrcProd!valorTotalProd) & "')"
    
        RecProd.Close
        
        RecOrcProd.MoveNext
    Loop
    RecOrcProd.Close
    
    Desconecta

    rptOrcamento.Show
End Sub
