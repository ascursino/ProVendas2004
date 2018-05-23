VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmProduto_Alt 
   Caption         =   "Alteração de Produtos"
   ClientHeight    =   4185
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
   Icon            =   "FrmProduto_Alt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   7320
   Begin VB.Frame FraArmacao 
      Height          =   3255
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   7095
      Begin VB.TextBox TxtPrVendaAtac 
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   6
         ToolTipText     =   "Preço de venda atacado"
         Top             =   2760
         Width           =   1455
      End
      Begin VB.ComboBox CboMoeda 
         Height          =   315
         ItemData        =   "FrmProduto_Alt.frx":0CCA
         Left            =   5520
         List            =   "FrmProduto_Alt.frx":0CD4
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Moeda"
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox TxtPrFabric 
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   4
         ToolTipText     =   "Preço do fornecedor"
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox TxtTipoProd 
         Height          =   285
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   2
         ToolTipText     =   "Tipo de produto"
         Top             =   1320
         Width           =   5055
      End
      Begin VB.TextBox TxtPrVendaUnit 
         Height          =   285
         Left            =   5520
         MaxLength       =   10
         TabIndex        =   5
         ToolTipText     =   "Preço de venda unitário"
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox TxtDescProd 
         Height          =   285
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   3
         ToolTipText     =   "Descrição do produto"
         Top             =   1800
         Width           =   5055
      End
      Begin VB.TextBox TxtProd 
         Height          =   285
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   1
         ToolTipText     =   "Nome do produto"
         Top             =   840
         Width           =   5055
      End
      Begin VB.ComboBox CboForn 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Fornecedor do produto"
         Top             =   360
         Width           =   4455
      End
      Begin VB.CommandButton CmdIncluirForn 
         Caption         =   "+"
         Height          =   255
         Left            =   6480
         TabIndex        =   10
         ToolTipText     =   "Adicionar fornecedor"
         Top             =   360
         Width           =   375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmProduto_Alt.frx":0CE5
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmProduto_Alt.frx":0D53
         TabIndex        =   14
         Top             =   1320
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   3720
         OleObjectBlob   =   "FrmProduto_Alt.frx":0DC5
         TabIndex        =   15
         Top             =   2760
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmProduto_Alt.frx":0E29
         TabIndex        =   16
         Top             =   840
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmProduto_Alt.frx":0E91
         TabIndex        =   17
         Top             =   1800
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   3720
         OleObjectBlob   =   "FrmProduto_Alt.frx":0F07
         TabIndex        =   18
         Top             =   2280
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmProduto_Alt.frx":0F87
         TabIndex        =   19
         Top             =   2280
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmProduto_Alt.frx":1001
         TabIndex        =   20
         Top             =   2760
         Width           =   1815
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
      TabIndex        =   11
      Top             =   3360
      Width           =   7095
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   1320
         OleObjectBlob   =   "FrmProduto_Alt.frx":1081
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
         Left            =   5760
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
         Left            =   4440
         TabIndex        =   8
         ToolTipText     =   "Efetuar inclusão"
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmProduto_Alt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String

Private Sub TxtPrFabric_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e ,===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 44 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPrFabric_LostFocus()
    If TxtPrFabric.Text <> "" Then
        TxtPrFabric.Text = Trim(Mid(FormataMoeda(TxtPrFabric.Text), 3))
    Else
        TxtPrFabric.Text = "0,00"
    End If
End Sub

Private Sub TxtProd_GotFocus()
    TxtProd.SelStart = 0
    TxtProd.SelLength = Len(TxtProd.Text)
End Sub

Private Sub TxtPrVendaAtac_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e ,===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 44 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPrVendaAtac_LostFocus()
    If TxtPrVendaAtac.Text <> "" Then
        TxtPrVendaAtac.Text = Trim(Mid(FormataMoeda(TxtPrVendaAtac.Text), 3))
    Else
        TxtPrVendaAtac.Text = "0,00"
    End If
End Sub

Private Sub TxtPrVendaUnit_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e ,===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 44 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPrVendaUnit_LostFocus()
    If TxtPrVendaUnit.Text <> "" Then
        TxtPrVendaUnit.Text = Trim(Mid(FormataMoeda(TxtPrVendaUnit.Text), 3))
    Else
        TxtPrVendaUnit.Text = "0,00"
    End If
End Sub

Private Sub TxtTipoProd_GotFocus()
    TxtTipoProd.SelStart = 0
    TxtTipoProd.SelLength = Len(TxtTipoProd.Text)
End Sub

Private Sub TxtDescProd_GotFocus()
    TxtDescProd.SelStart = 0
    TxtDescProd.SelLength = Len(TxtDescProd.Text)
End Sub

Private Sub TxtPrFabric_GotFocus()
    TxtPrFabric.SelStart = 0
    TxtPrFabric.SelLength = Len(TxtPrFabric.Text)
End Sub

Private Sub TxtPrVendaUnit_GotFocus()
    TxtPrVendaUnit.SelStart = 0
    TxtPrVendaUnit.SelLength = Len(TxtPrVendaUnit.Text)
End Sub

Private Sub TxtPrVendaAtac_GotFocus()
    TxtPrVendaAtac.SelStart = 0
    TxtPrVendaAtac.SelLength = Len(TxtPrVendaAtac.Text)
End Sub

Private Sub CmdFechar_Click()
    Unload Me
    
    MDIPrincipal.Enabled = True
    MDIPrincipal.WindowState = 2
End Sub

Private Sub CmdIncluirForn_Click()
    VGStrIncluirProd = "altprod"
    FrmFornecedor_Inc.Show
End Sub

Private Sub CmdOK_Click()
    If TxtProd.Text = "" Or TxtPrVendaUnit.Text = "" Then
        VPStrBox = MsgBox("Preencha pelo menos os campos principais." & Chr(13) & "(Produto e Preço venda (unit.))", vbInformation, "Pró Vendas 2004 - Informação")
    Else
        If TxtPrFabric.Text = "" Then
            TxtPrFabric.Text = "0,00"
        End If
        
        If TxtPrVendaUnit.Text = "" Then
            TxtPrVendaUnit.Text = "0,00"
        End If
        
        If TxtPrVendaAtac.Text = "" Or TxtPrVendaAtac.Text = "0,00" Then
            TxtPrVendaAtac.Text = TxtPrVendaUnit.Text
        End If
    
        Dim RecProd As New ADODB.Recordset
        
        Conecta
        
        StrSql = "Select * from tb_Produto where CodProd=" & VGIntCodProd
        RecProd.Open StrSql, vgCon, 1, 3
        
        If CboForn.Text <> "" Then
            RecProd("CodForn") = Trim(Mid(CboForn.Text, Len(CboForn.Text) - 10))
        Else
            RecProd("CodForn") = 0
        End If
        RecProd("NomeProd") = TxtProd.Text
        RecProd("TipoProd") = TxtTipoProd.Text
        RecProd("DescProd") = TxtDescProd.Text
        RecProd("PrecoFabric") = TxtPrFabric.Text
        RecProd("PrecoVendaUnit") = TxtPrVendaUnit.Text
        RecProd("PrecoVendaAtac") = TxtPrVendaAtac.Text
        RecProd("Moeda") = CboMoeda.Text
        RecProd.Update
        
        Desconecta
        
        VPStrBox = MsgBox("Alteração efetuada.", vbInformation, "Pró Vendas 2004 - Informação")
        
        FrmPrincipal.CmdPesqProd.Value = True
    
        MDIPrincipal.Enabled = True
        
        Unload Me
    End If
End Sub

Private Sub Form_Resize()
  FrmProduto_Alt.Left = (MDIPrincipal.Width / 2) - (FrmProduto_Alt.Width / 2)
  FrmProduto_Alt.Top = (MDIPrincipal.Height / 3) - (FrmProduto_Alt.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 4695
    Width = 7440
    
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    Call MontaCboForn
    
    Dim RecProd As New ADODB.Recordset
    Dim RecForn As New ADODB.Recordset
    
    Conecta
    
    StrSql = "Select * from tb_Produto where CodProd=" & VGIntCodProd
    RecProd.Open StrSql, vgCon, 1, 3
    
    StrSql = "Select CodForn,Nome from tb_Fornecedor where CodForn=" & RecProd!CodForn
    RecForn.Open StrSql, vgCon, 1, 3
    
    If Not RecForn.EOF Then
        CboForn.Text = RecForn!nome & "                                                                                                 " & RecForn!CodForn
    End If
    TxtProd.Text = VerificaNulo(RecProd!nomeprod)
    TxtTipoProd.Text = VerificaNulo(RecProd!tipoprod)
    TxtDescProd.Text = VerificaNulo(RecProd!descprod)
    TxtPrFabric.Text = VerificaNulo(RecProd!precofabric)
    TxtPrVendaUnit.Text = VerificaNulo(RecProd!precovendaunit)
    TxtPrVendaAtac.Text = VerificaNulo(RecProd!precovendaatac)
    CboMoeda.Text = RecProd!moeda
    
    Desconecta
    
    MDIPrincipal.Enabled = False
End Sub

Sub MontaCboForn()
    Dim RecCbo As New ADODB.Recordset
    
    CboForn.Clear
    
    Conecta
    
    StrSql = "Select CodForn,Nome from tb_Fornecedor"
    RecCbo.Open StrSql, vgCon, 1, 3
    
    CboForn.AddItem ("")
    Do While Not RecCbo.EOF
        CboForn.AddItem (RecCbo!nome & "                                                                                                 " & RecCbo!CodForn)
        RecCbo.MoveNext
    Loop
    
    RecCbo.Close
    
    Desconecta

End Sub

