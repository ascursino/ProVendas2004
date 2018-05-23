VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form FrmEstoque_Alerta 
   Caption         =   "Alerta de Estoque"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10080
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
   Icon            =   "FrmEstoque_Alerta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   10080
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
      Top             =   3000
      Width           =   9855
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
         Left            =   7200
         TabIndex        =   1
         ToolTipText     =   "Imprimir alerta de estoque"
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   0
         OleObjectBlob   =   "FrmEstoque_Alerta.frx":0CCA
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
         Left            =   8520
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
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9855
      Begin FPSpread.vaSpread GridAlerta 
         Height          =   2055
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   9615
         _Version        =   393216
         _ExtentX        =   16960
         _ExtentY        =   3625
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "FrmEstoque_Alerta.frx":0EFE
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmEstoque_Alerta.frx":12AE
         TabIndex        =   5
         Top             =   240
         Width           =   4095
      End
   End
End
Attribute VB_Name = "FrmEstoque_Alerta"
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

    Dim prod As String
    Dim qtdeprod As String
    Dim qtdemin As String
    Dim forn As String
    Dim tel As String

    Dim VLStrLinha As String

    VLStrLinha = 1

    Conecta

    Do While VLStrLinha <= GridAlerta.MaxRows

        GridAlerta.Col = 1
        GridAlerta.Row = VLStrLinha
        prod = GridAlerta.Text

        GridAlerta.Col = 2
        GridAlerta.Row = VLStrLinha
        qtdeprod = GridAlerta.Text

        GridAlerta.Col = 3
        GridAlerta.Row = VLStrLinha
        qtdemin = GridAlerta.Text

        GridAlerta.Col = 4
        GridAlerta.Row = VLStrLinha
        forn = GridAlerta.Text

        GridAlerta.Col = 5
        GridAlerta.Row = VLStrLinha
        tel = GridAlerta.Text

        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05) " & _
        "VALUES ('" & prod & "','" & qtdeprod & "','" & qtdemin & "','" & forn & "','" & tel & "')"

        VLStrLinha = VLStrLinha + 1
    Loop

    Desconecta

    rptEstoqueAlerta.Show

End Sub

Private Sub Form_Load()
    Height = 4350
    Width = 10200
    
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    MDIPrincipal.Enabled = False
        
    Conecta
    
    Dim RecAlert As New ADODB.Recordset
    Dim RecForn As New ADODB.Recordset
    Dim VLStrTel1 As String
    Dim VLStrTel2 As String
    
    StrSql = "SELECT E.QtdeMin,E.QtdeProd,P.NomeProd,P.CodForn " & _
             "FROM tb_Estoque as E,tb_Produto as P " & _
             "WHERE E.CodProd=P.CodProd AND E.QtdeProd <= E.QtdeMin " & _
             "ORDER BY P.NomeProd"
    RecAlert.Open StrSql, vgCon, 1, 3

    If Not RecAlert.EOF Then
        'Dim VLIntCodEst As Long
        Dim VLIntLinha As Long
    
        VLIntLinha = 1
        GridAlerta.MaxRows = VLIntLinha

        Do While Not RecAlert.EOF
            GridAlerta.Row = VLIntLinha
            GridAlerta.Lock = True

            'Produto
            GridAlerta.Col = 1
            GridAlerta.Text = VerificaNulo(RecAlert!nomeprod)
            GridAlerta.Lock = True

            'Qtde em estoque
            GridAlerta.Col = 2
            GridAlerta.Text = VerificaNulo(RecAlert!qtdeprod)
            GridAlerta.Lock = True

            'Qtde mínima
            GridAlerta.Col = 3
            GridAlerta.Text = VerificaNulo(RecAlert!qtdemin)
            GridAlerta.Lock = True

            'Fornecedor
            StrSql = "SELECT Nome,Telefone1,Telefone2 FROM tb_fornecedor WHERE CodForn=" & RecAlert!CodForn
            RecForn.Open StrSql, vgCon, 1, 3
            
            GridAlerta.Col = 4
            If Not RecForn.EOF Then
                GridAlerta.Text = VerificaNulo(RecForn!nome)
            Else
                GridAlerta.Text = ""
            End If
            GridAlerta.Lock = True

            'Telefone
            GridAlerta.Col = 5
            If Not RecForn.EOF Then
                If RecForn!telefone1 <> "" And IsNull(RecForn!telefone1) = False Then
                    VLStrTel1 = RecForn!telefone1
                Else
                    VLStrTel1 = ""
                End If
                
                If RecForn!telefone2 <> "" And IsNull(RecForn!telefone2) = False Then
                    VLStrTel2 = RecForn!telefone2
                Else
                    VLStrTel2 = ""
                End If
                
                If VLStrTel1 = "" And VLStrTel2 = "" Then
                    GridAlerta.Text = ""
                    
                ElseIf VLStrTel1 <> "" And VLStrTel2 <> "" Then
                    GridAlerta.Text = VLStrTel1 & " / " & VLStrTel2
                    
                ElseIf VLStrTel1 <> "" And VLStrTel2 = "" Then
                    GridAlerta.Text = VLStrTel1
                
                ElseIf VLStrTel1 = "" And VLStrTel2 <> "" Then
                    GridAlerta.Text = VLStrTel2
                    
                End If
                
            Else
                GridAlerta.Text = ""
            End If
            GridAlerta.Lock = True
            
            RecForn.Close
            
            VLIntLinha = VLIntLinha + 1

            GridAlerta.MaxRows = GridAlerta.MaxRows + 1
            RecAlert.MoveNext
        Loop
        
        GridAlerta.MaxRows = GridAlerta.MaxRows - 1
        
        CmdImprimir.Enabled = True
        
    Else
        CmdImprimir.Enabled = False
    End If
    
    Desconecta

End Sub

Private Sub Form_Resize()
  FrmEstoque_Alerta.Left = (MDIPrincipal.Width / 2) - (FrmEstoque_Alerta.Width / 2)
  FrmEstoque_Alerta.Top = (MDIPrincipal.Height / 3) - (FrmEstoque_Alerta.Height / 3)
End Sub
