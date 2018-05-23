VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "spr32x60.ocx"
Begin VB.Form FrmParcela_Inc 
   Caption         =   "Inclusão de parcelas do crediário"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9090
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
   Icon            =   "FrmParcela_Inc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   9090
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
      Top             =   3360
      Width           =   8895
      Begin VB.CommandButton CmdOK 
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
         Left            =   6240
         TabIndex        =   1
         ToolTipText     =   "Efetuar inclusão"
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   960
         OleObjectBlob   =   "FrmParcela_Inc.frx":0CCA
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
         Left            =   7560
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
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   8895
      Begin FPSpread.vaSpread GridParcela 
         Height          =   2895
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   8655
         _Version        =   393216
         _ExtentX        =   15266
         _ExtentY        =   5106
         _StockProps     =   64
         ColHeaderDisplay=   0
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
         Protect         =   0   'False
         RowHeaderDisplay=   2
         ScrollBarExtMode=   -1  'True
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "FrmParcela_Inc.frx":0EFE
         UserResize      =   1
      End
   End
End
Attribute VB_Name = "FrmParcela_Inc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdFechar_Click()
    Unload Me
   
    FrmVenda_Inc.Enabled = True
End Sub

Private Sub CmdOK_Click()
    Dim VLIntLinha As Integer
    Dim VLIntLinhaMax As Integer
    Dim VLStrCheque As String
    Dim VLStrDigito As String

    VLIntLinha = 1
    VGStrBanco = ""
    VLStrCheque = ""
    VLStrDigito = ""
    VGStrChequeDigito = ""
    VGStrData = ""
    VGStrValor = ""
    
    Do While VLIntLinha <= GridParcela.MaxRows
                
        GridParcela.Row = VLIntLinha
        
        'Banco
        GridParcela.Col = 1
        If GridParcela.Text <> "" Then
            VGStrBanco = VGStrBanco & " # " & Trim(GridParcela.Text)
        Else
            VGStrBanco = VGStrBanco & " # 0"
        End If
          
        'Cheque
        GridParcela.Col = 2
        If GridParcela.Text <> "" Then
            VLStrCheque = Trim(GridParcela.Text)
        Else
            VLStrCheque = ""
        End If
        
        'Dígito do cheque
        GridParcela.Col = 3
        If GridParcela.Text <> "" Then
            VLStrDigito = Trim(GridParcela.Text)
        Else
            VLStrDigito = ""
        End If
        
        'juntando cheque com dígito
        If VLStrCheque <> "" And VLStrDigito <> "" Then
            VGStrChequeDigito = VGStrChequeDigito & " # " & VLStrCheque & "-" & VLStrDigito
        ElseIf VLStrCheque <> "" And VLStrDigito = "" Then
            VGStrChequeDigito = VGStrChequeDigito & " # " & VLStrCheque
        ElseIf (VLStrCheque = "" And VLStrDigito <> "") Or (VLStrCheque = "" And VLStrDigito = "") Then
            VGStrChequeDigito = VGStrChequeDigito & " # 0"
        End If
        
        'Data
        GridParcela.Col = 4
        If GridParcela.Text <> "" Then
            VGStrData = VGStrData & " # " & Trim(GridParcela.Text)
        Else
            VGStrData = VGStrData & " # 0"
        End If
        
        'Valor
        GridParcela.Col = 5
        If GridParcela.Text <> "" Then
            VGStrValor = VGStrValor & " # " & Trim(GridParcela.Text)
        Else
            VGStrValor = VGStrValor & " # 0"
        End If
        
        VLIntLinha = VLIntLinha + 1
    Loop
    
    VGStrBanco = VGStrBanco & " #"
    VGStrChequeDigito = VGStrChequeDigito & " #"
    VGStrData = VGStrData & " #"
    VGStrValor = VGStrValor & " #"
    
    VGStrParcelaCheque = "sim"
    
    Unload Me
    
    FrmVenda_Inc.Enabled = True
    
End Sub

Private Sub Form_Resize()
  FrmParcela_Inc.Left = (MDIPrincipal.Width / 2) - (FrmParcela_Inc.Width / 2)
  FrmParcela_Inc.Top = (MDIPrincipal.Height / 3) - (FrmParcela_Inc.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 4695
    Width = 9210
    
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    FrmVenda_Inc.Enabled = False
    
    Dim VLIntLinha As Integer
    Dim VLIntLinhaMax As Integer
    Dim VLStrValorParc As String
    Dim VLStrData As String
    
    VLIntLinhaMax = FrmVenda_Inc.CboPrazoChqParc.Text
    VLIntLinha = 1
    
    Do While VLIntLinha <= VLIntLinhaMax
                
        GridParcela.Row = VLIntLinha
        GridParcela.Col = 0
        GridParcela.Text = FormataNum(VLIntLinha) & "ª parcela"
        
        VLIntLinha = VLIntLinha + 1
        GridParcela.MaxRows = GridParcela.MaxRows + 1
    Loop
    GridParcela.MaxRows = GridParcela.MaxRows - 1
    
    If VGStrBanco = "" And VGStrChequeDigito = "" And VGStrData = "" And VGStrValor = "" Then
        VLIntLinha = 1
        VLStrValorParc = Mid(FrmVenda_Inc.LblParcChq.Caption, InStr(FrmVenda_Inc.LblParcChq.Caption, "R$"))
        VLStrData = Date
        
        Do While VLIntLinha <= VLIntLinhaMax
            VLStrData = DateSerial(Year(VLStrData), Month(VLStrData), Day(VLStrData) + 30)
            
            GridParcela.Row = VLIntLinha
            GridParcela.Col = 4
            GridParcela.Text = FormataData(VLStrData)
            
            GridParcela.Row = VLIntLinha
            GridParcela.Col = 5
            GridParcela.Text = VLStrValorParc
            
            VLIntLinha = VLIntLinha + 1
        Loop
    Else
        Dim VLStrBancoTemp As String
        Dim VLStrChequeDigitoTemp As String
        Dim VLStrChequeTemp As String
        Dim VLStrDigitoTemp As String
        Dim VLStrDataTemp As String
        Dim VLStrValorTemp As String
        
        VLStrBancoTemp = VGStrBanco
        VLStrChequeDigitoTemp = VGStrChequeDigito
        VLStrDataTemp = VGStrData
        VLStrValorTemp = VGStrValor
        
        VLIntLinha = 1
        
        Do While VLIntLinha <= VLIntLinhaMax
            GridParcela.Row = VLIntLinha
            
            'Banco
            GridParcela.Col = 1
            If VLStrBancoTemp <> "" And VLStrBancoTemp <> "#" Then
                VLStrBancoTemp = Trim(Mid(VLStrBancoTemp, InStr(VLStrBancoTemp, "#") + 1))
                
                GridParcela.Text = Trim(Mid(VLStrBancoTemp, 1, InStr(VLStrBancoTemp, "#") - 1))
                
                VLStrBancoTemp = Trim(Mid(VLStrBancoTemp, InStr(VLStrBancoTemp, "#")))
            Else
                GridParcela.Text = ""
            End If
              
            'Cheque e dígito
            If VLStrChequeDigitoTemp <> "" And VLStrChequeDigitoTemp <> "#" Then
                VLStrChequeDigitoTemp = Trim(Mid(VLStrChequeDigitoTemp, InStr(VLStrChequeDigitoTemp, "#") + 1))
                VLStrChequeTemp = Trim(Mid(VLStrChequeDigitoTemp, 1, InStr(VLStrChequeDigitoTemp, "#") - 1))
                
                If InStr(VLStrChequeTemp, "-") <> 0 Then
                    GridParcela.Col = 2
                    GridParcela.Text = Trim(Mid(VLStrChequeTemp, 1, InStr(VLStrChequeTemp, "-") - 1))
                    
                    GridParcela.Col = 3
                    GridParcela.Text = Trim(Mid(VLStrChequeTemp, InStr(VLStrChequeTemp, "-") + 1))
                Else
                    GridParcela.Col = 2
                    GridParcela.Text = VLStrChequeTemp
                    
                    GridParcela.Col = 3
                    GridParcela.Text = ""
                End If
                
                VLStrChequeDigitoTemp = Trim(Mid(VLStrChequeDigitoTemp, InStr(VLStrChequeDigitoTemp, "#")))
            Else
                GridParcela.Col = 2
                GridParcela.Text = ""
                
                GridParcela.Col = 3
                GridParcela.Text = ""
            End If
            
            'Data
            GridParcela.Col = 4
            If VLStrDataTemp <> "" And VLStrDataTemp <> "#" Then
                VLStrDataTemp = Trim(Mid(VLStrDataTemp, InStr(VLStrDataTemp, "#") + 1))
                
                GridParcela.Text = Trim(Mid(VLStrDataTemp, 1, InStr(VLStrDataTemp, "#") - 1))
                
                VLStrDataTemp = Trim(Mid(VLStrDataTemp, InStr(VLStrDataTemp, "#")))
            Else
                GridParcela.Text = ""
            End If
            
            'Valor
            GridParcela.Col = 5
            If VLStrValorTemp <> "" And VLStrValorTemp <> "#" Then
                VLStrValorTemp = Trim(Mid(VLStrValorTemp, InStr(VLStrValorTemp, "#") + 1))
                
                GridParcela.Text = Trim(Mid(VLStrValorTemp, 1, InStr(VLStrValorTemp, "#") - 1))
                
                VLStrValorTemp = Trim(Mid(VLStrValorTemp, InStr(VLStrValorTemp, "#")))
            Else
                GridParcela.Text = ""
            End If
            
            VLIntLinha = VLIntLinha + 1
        Loop
    End If
End Sub

