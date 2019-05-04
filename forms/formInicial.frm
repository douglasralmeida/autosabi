VERSION 5.00
Begin VB.Form formInicial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Automatizador do SABI"
   ClientHeight    =   6750
   ClientLeft      =   150
   ClientTop       =   330
   ClientWidth     =   8055
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "formInicial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox painelErro 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1335
      ScaleWidth      =   7695
      TabIndex        =   35
      Top             =   5280
      Width           =   7695
      Begin VB.CommandButton btoFecharErro 
         Caption         =   "Fechar"
         Height          =   360
         Left            =   6480
         TabIndex        =   37
         ToolTipText     =   " Fechar aplicativo "
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label txtErro 
         Caption         =   "Label3"
         Height          =   255
         Left            =   1080
         TabIndex        =   36
         Top             =   360
         Width           =   6375
      End
   End
   Begin VB.PictureBox painelStatus 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   120
      ScaleHeight     =   975
      ScaleWidth      =   7695
      TabIndex        =   32
      Top             =   4200
      Width           =   7695
      Begin VB.Label txtStatusAguarda 
         Caption         =   "Label3"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   7335
      End
      Begin VB.Label txtStatus 
         Caption         =   "Label3"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   7335
      End
   End
   Begin VB.PictureBox pctCopiaPartedaTelaCPF 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4320
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   127
      TabIndex        =   23
      Top             =   3840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox pctImpressora 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   7800
      Picture         =   "formInicial.frx":08CA
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   22
      Top             =   600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pctFundo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   120
      ScaleHeight     =   3855
      ScaleWidth      =   8175
      TabIndex        =   5
      Top             =   0
      Width           =   8175
      Begin VB.Frame grupoOrdem 
         Caption         =   "Requerimentos ordenados por"
         ForeColor       =   &H00404040&
         Height          =   855
         Left            =   0
         TabIndex        =   0
         Top             =   1800
         Width           =   7815
         Begin VB.CommandButton btoFechar 
            Cancel          =   -1  'True
            Caption         =   "Fechar"
            Height          =   370
            Left            =   6480
            TabIndex        =   4
            Top             =   280
            Width           =   1200
         End
         Begin VB.CommandButton btoIniciar 
            Caption         =   "&Processar"
            Default         =   -1  'True
            Height          =   375
            Left            =   5160
            TabIndex        =   3
            Top             =   280
            Width           =   1200
         End
         Begin VB.OptionButton optOrdem 
            Caption         =   "&Nome do Periciando"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   1
            ToolTipText     =   "Apresenta os requerimentos ordenados por nome do periciando"
            Top             =   400
            Value           =   -1  'True
            Width           =   2295
         End
         Begin VB.OptionButton optOrdem 
            Caption         =   "&Hora da Perícia"
            Height          =   255
            Index           =   1
            Left            =   2640
            TabIndex        =   2
            ToolTipText     =   "Apresenta os requerimentos ordenados pelo horario da pericia"
            Top             =   400
            Width           =   2175
         End
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   5040
         Top             =   120
      End
      Begin VB.Timer timerAbrirSabi 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   4560
         Top             =   120
      End
      Begin VB.Frame fraImprime 
         Caption         =   "Imprimir 2ª Via da Marcação de Exame"
         ForeColor       =   &H00404040&
         Height          =   1005
         Left            =   0
         TabIndex        =   20
         Top             =   2760
         Visible         =   0   'False
         Width           =   7815
         Begin VB.CommandButton btoFechar2 
            Caption         =   "Fechar"
            Height          =   360
            Left            =   6480
            TabIndex        =   31
            ToolTipText     =   " Fechar aplicativo "
            Top             =   360
            Width           =   1200
         End
         Begin VB.CheckBox chcPP 
            Caption         =   "PP"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2880
            TabIndex        =   29
            ToolTipText     =   " marcar para imprimir os exames de PP "
            Top             =   680
            Width           =   735
         End
         Begin VB.CheckBox chcExameInicial 
            Caption         =   "Exame Inicial"
            Height          =   240
            Left            =   360
            TabIndex        =   28
            ToolTipText     =   " marcar para imprimir os exames iniciais "
            Top             =   680
            Width           =   1695
         End
         Begin VB.TextBox editMarcacaoPara 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   330
            Left            =   2880
            TabIndex        =   25
            Text            =   "1"
            ToolTipText     =   " fixar o final da sequ?ncia de impress?o "
            Top             =   240
            Width           =   500
         End
         Begin VB.TextBox editMarcacaoDe 
            Height          =   330
            Left            =   1800
            TabIndex        =   24
            Text            =   "1"
            ToolTipText     =   " fixar o ?nicio da sequ?ncia de impress?o "
            Top             =   240
            Width           =   500
         End
         Begin VB.CommandButton btoConfirmar 
            Caption         =   "&Confirmar"
            Height          =   360
            Left            =   5160
            TabIndex        =   21
            ToolTipText     =   " Confirmar a sequ?ncia e os tipos de exames e inciar a opera??o de impress?o "
            Top             =   360
            Width           =   1200
         End
         Begin VB.Image parabaixo 
            Height          =   240
            Left            =   4200
            Picture         =   "formInicial.frx":0BAC
            Stretch         =   -1  'True
            ToolTipText     =   " mover a lista de requerimentos para baixo "
            Top             =   600
            Width           =   240
         End
         Begin VB.Image paracima 
            Height          =   240
            Left            =   4200
            Picture         =   "formInicial.frx":0FEE
            Stretch         =   -1  'True
            ToolTipText     =   " mover a lista de requerimentos para cima "
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "a "
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2280
            TabIndex        =   27
            Top             =   300
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Sequencia: De"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   360
            TabIndex        =   26
            Top             =   300
            Width           =   1335
         End
      End
      Begin VB.PictureBox pctProgressoFundo 
         Appearance      =   0  'Flat
         FillColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   7440
         ScaleHeight     =   225
         ScaleWidth      =   465
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   495
         Begin VB.PictureBox pctProgresso 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            FillColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   600
            Left            =   -30
            ScaleHeight     =   570
            ScaleWidth      =   165
            TabIndex        =   19
            Top             =   -30
            Width           =   200
         End
      End
      Begin VB.PictureBox pctFundoCopias 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   3480
         ScaleHeight     =   735
         ScaleWidth      =   2535
         TabIndex        =   14
         Top             =   600
         Visible         =   0   'False
         Width           =   2535
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1455
            Left            =   1080
            ScaleHeight     =   97
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   129
            TabIndex        =   15
            Top             =   120
            Width           =   1935
         End
      End
      Begin VB.PictureBox pctApresentaPartedaTelaCopiada 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5520
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   121
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ListBox listaRequerimentos 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         ItemData        =   "formInicial.frx":1430
         Left            =   240
         List            =   "formInicial.frx":1437
         TabIndex        =   12
         Top             =   1080
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.PictureBox pctCopiaPartedaTela 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2880
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   121
         TabIndex        =   10
         Top             =   1080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ListBox listaClassificar 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblLocaleData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "local e data"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   5880
         TabIndex        =   30
         Top             =   1440
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Image imageIcone 
         Height          =   480
         Left            =   80
         Picture         =   "formInicial.frx":144F
         Top             =   80
         Width           =   480
      End
      Begin VB.Label lblRelogio 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 00 "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   225
         Left            =   7560
         TabIndex        =   11
         Top             =   840
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lbNomePrograma 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Automatizador do SABI"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   9
         Top             =   120
         Width           =   2925
      End
      Begin VB.Label lbVersao 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Compilado em 05-04-2019"
         ForeColor       =   &H00808080&
         Height          =   225
         Left            =   720
         TabIndex        =   8
         Top             =   480
         Width           =   2100
      End
      Begin VB.Label lblImprimir 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   " Imprimir  "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   240
         Left            =   7080
         TabIndex        =   7
         Top             =   1200
         Visible         =   0   'False
         Width           =   840
      End
   End
   Begin VB.PictureBox pctDigitos 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   6960
      Picture         =   "formInicial.frx":1D19
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   16
      Top             =   3840
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.PictureBox pctEsteRequerimento 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   5280
      Picture         =   "formInicial.frx":2BFB
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   178
      TabIndex        =   17
      Top             =   3240
      Visible         =   0   'False
      Width           =   2670
   End
End
Attribute VB_Name = "formInicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Sabi As classeSabi

Dim MenuName As New Collection
Dim MenuHandle As New Collection
Dim lHwnd As Long
Dim imprimereq As Boolean
Dim modoImprime As String
Dim LocalY As Long
Dim LocalCopiar As Boolean
Dim requerimentomostrado As Long
Dim contadorTimer As Long
Dim mtempo2 As Long
Dim deslocalista As Long
    
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, wParam As Any, lParam As String) As Long

Public Sub excluirArquivosTemp()
  Dim memo As String
  
  'apaga todos bmp de datas anteriores a atual
  memo = Dir(GlobalPastadeTrabalho & "\" & "*.bmp")
  While memo <> ""
    If Mid(memo, 1, 8) < Format(Date, "yyyymmdd") Then Kill GlobalPastadeTrabalho & "\" & memo
    memo = Dir()
  Wend
  
  'apaga todos txt de datas anteriores a atual
  memo = Dir(GlobalPastadeTrabalho & "\" & "*.txt")
  While memo <> ""
    If Mid(memo, 1, 8) < Format(Date, "yyyymmdd") Then Kill GlobalPastadeTrabalho & "\" & memo
    memo = Dir()
  Wend
  memo = Dir(GlobalPastadeTrabalho & "\*" & ".txt")
  Do While memo <> ""
    If IsDate(Mid(memo, 7, 2) & "/" & Mid(memo, 5, 2) & "/" & Mid(memo, 1, 4)) Then
      Exit Do
    End If
    memo = Dir()
  Loop
End Sub

Sub exibirRelatorioFinal()
  On Error Resume Next
  Dim conta As Long
  Dim res As String
  Dim nome As String
  Dim segundos As Long
  Dim minutos As Long
  Dim idArquivo As Integer
  
  arquivoRelatorio = GlobalPastadeTrabalho & "\relatoriofinal.txt"
  
  listaRequerimentos.Enabled = True
  segundos = Int((GetTickCount - GlobalInicio) / 1000)
  minutos = Int(segundos / 60)
  segundos = segundos - minutos * 60
  lblRelogio.Caption = " " & minutos & ":" & Format(segundos, "00") & " "
  texto = "O Controle Operacional foi fechado por medida de segurança" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
  texto = texto & lblLocaleData.Caption & Chr(13) & Chr(10)
  texto = texto & acertarLarguraColuna("123", "Seq") & Chr(9) & acertarLarguraColuna("123456789", "Requerimento") & Chr(9) & acertarLarguraColuna("12345678901", "CPF") & Chr(9) & acertarLarguraColuna("INICIAL", "Tipo") & Chr(9) & acertarLarguraColuna("INDEFERIDO", "Status") & Chr(9) & acertarLarguraColuna("12345678901", "NIT") & Chr(9) & acertarLarguraColuna("NÃO", "IMPRESSO") & Chr(9) & acertarLarguraColuna("JOSE GERALDO DA COSTA", "Segurado") & Chr(9) & "Crítica"
  For conta = 1 To QuantidadedeRequerimentos
    texto = texto & Chr(13) & Chr(10)
    texto = texto & acertarLarguraColuna("123", Format(conta, "000"))
    texto = texto & Chr(9) & acertarLarguraColuna("123456789", requerimentos(conta).Número)
    texto = texto & Chr(9) & acertarLarguraColuna("12345678901", requerimentos(conta).CPF)
    texto = texto & Chr(9) & acertarLarguraColuna("INICIAL", requerimentos(conta).Tipo)
    texto = texto & Chr(9) & acertarLarguraColuna("INDEFERIDO", requerimentos(conta).Status)
    texto = texto & Chr(9) & acertarLarguraColuna("12345678901", requerimentos(conta).nit)
    texto = texto & Chr(9) & acertarLarguraColuna("NÃO", requerimentos(conta).Impresso)
    texto = texto & Chr(9) & acertarLarguraColuna("JOSE GERALDO DA COSTA", requerimentos(conta).Segurado)
    texto = texto & Chr(9) & Trim(requerimentos(conta).Crítica)
  Next conta

  idArquivo = FreeFile
  
  ' Excluir arquivo existente
  excluirArquivo arquivoRelatorio
  
  ' Abre o arquivo
  Open arquivoRelatorio For Output As #idArquivo
    Print #idArquivo, texto
  
  'Fecha o arquivo
  Close #idArquivo
  
  abrirArquivo (arquivoRelatorio)
  
  GlobalRelatorioPronto = True
End Sub

Public Sub exibirStatus(Status As String, segundos As Integer)
  pctFundo.Visible = False
  txtStatus.Caption = Status
  txtStatusAguarda.Caption = "Aguardando " & segundos & " segundos..."
  painelStatus.Top = 0
  painelStatus.Left = 0
  painelStatus.Visible = True
  contadorTimer = segundos
  timerAbrirSabi.Enabled = True
End Sub

Sub obterDadosRegistro()
  Dim hkey As Long
  Dim imprimirExames As String
  Dim valorIniciais As Long
  Dim valorPP As Long
  Dim imprimirOrdem As String
  Dim valorOrdem As Boolean
  Dim valorTempoEspera As Long
  
  valorIniciais = 1
  valorPP = 0
  valorOrdem = False
  valorTempoEspera = 3
  'If abrirRegChave(hkey) Then
  '  imprimirExames = lerRegValor(hkey, "ImprimirExames", "INICIAIS")
  '  If imprimirExames = "NENHUM" Then
  '    valorIniciais = 0
  '    valorPP = 0
  '  ElseIf imprimirExames = "TODOS" Then
  '     valorIniciais = 1
  '     valorPP = 1
  '  ElseIf imprimirExames = "INICIAIS" Then
  '    valorIniciais = 1
  '    valorPP = 0
  '  Else 'PP
  '    valorIniciais = 0
  '    valorPP = 1
  '  End If
  '  imprimirOrdem = lerRegValor(hkey, "ImprimirOrdem", "HORA")
  '  If imprimirOrdem = "HORA" Then
  '    valorOrdem = True
  '  Else
  '    valorOrdem = False
  '  valorTempoEspera = Val(lerRegValor(hkey, "TempodeEsperadaResposta", "3"))
  '  If valorTempoEspera < 3 Or valorTempoEspera > 10 Then valorTempoEspera = 3
  'End If
  '
  
  formInicial.chcExameInicial.Value = valorIniciais
  formInicial.chcPP.Value = valorPP
  formInicial.optOrdem(1).Value = valorOrdem
  GlobalTempodeEspera = valorTempoEspera
End Sub
    
Sub obterDadosSistema()
  On Error Resume Next 'voltar
  Dim lpBuff As String * 25
  Dim ret As Long

  'Get the user name minus any trailing spaces found in the name.
  ret = GetUserName(lpBuff, 25)
  GlobalUserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
  GlobalAreadeTrabalho = getPastaEspecial(CSIDL_DESKTOP)
  GlobalPastadeTrabalho = getPastaEspecial(CSIDL_LOCAL_APPDATA) & "\" & NomeAplicacao
End Sub

Public Sub redimensionarForm(topo As Integer, altura As Integer)
  'Em VB as medidas da tela sao em twips, e nao pixels.
  formInicial.Top = Screen.Height + topo
  formInicial.Height = altura
End Sub

Private Function testarSabiAberto() As Boolean
  If Sabi Is Nothing Then Set Sabi = New classeSabi
  testarSabiAberto = Sabi.estaAberto
End Function

Private Function processar() As Boolean
  Sabi.prepararAmbiente
  Sabi.abrirJanelaCarteira
  Sabi.abrirJanelaAgendamentos
  Sabi.processarDataAgenda
  Sabi.fecharListaAgendamentos
  Sabi.exportarAgendamentos
  Sabi.definirArquivoAgendamento
  Sabi.fecharJanelaCrystalReport
  Sabi.processarRequerimentos
  Sabi.relatorioFinal
  Sabi.exibirAgendamentos
  desenharLista
End Function

Sub mostratela()
  Dim RtnValue
  Dim win As Long
  Dim desloca As Long
  Dim esquerda, altura, largura, dimensao As Long
  If Val(GlobalRequerimentos(GlobalIDRequerimento).Número) = 0 Or (requerimentomostrado = Val(GlobalRequerimentos(GlobalIDRequerimento).Número)) Then Exit Sub
  If Val(GlobalRequerimentos(GlobalIDRequerimento).sequencia) > 40 Then 'numero de linhas
    desloca = ((Val(GlobalRequerimentos(GlobalIDRequerimento).sequencia) - 40) * 240) / 15
  Else
    desloca = 0
  End If
  esquerda = 760
  altura = 0
  largura = Picture1.Width
  dimensao = Picture1.Height
  win = GlobalhMDIClient
  Picture1.Refresh
  pctEsteRequerimento.Refresh
  requerimentomostrado = Val(GlobalRequerimentos(GlobalIDRequerimento).Número)
  RtnValue = BitBlt(GetDC(win), CLng(esquerda), CLng(-desloca), CLng(largura), CLng(dimensao), Picture1.hDC, CLng(0), CLng(0), SRCCOPY)
End Sub
    
Public Function CapturaNumeroDetalhes() As String
  On Error Resume Next
  Dim TopRequerimento As Long
  Dim LeftRequerimento As Long
  Dim algarismo As Long
  Dim soma As Long
  Dim indice As Long
  Dim digito As Long
  Dim letra As String
  Dim Requerimento As String
  Dim deslocamento As Long
  Dim Digitos As Long

  Requerimento = ""
  TopRequerimento = 2
  deslocamento = 1
  Digitos = 11
  LeftRequerimento = deslocamento
  soma = 0
  algarismo = 0
  For digito = 0 To Digitos - 1
    soma = 0
    If pctCopiaPartedaTelaCPF.Point(algarismo + LeftRequerimento, TopRequerimento + 2) = 0 Then
      soma = 1
    End If
    If pctCopiaPartedaTelaCPF.Point(algarismo + LeftRequerimento, TopRequerimento + 6) = 0 Then
      soma = soma + 2
    End If
    If pctCopiaPartedaTelaCPF.Point(algarismo + LeftRequerimento + 2, TopRequerimento) = 0 Then
      soma = soma + 4
    End If
    If pctCopiaPartedaTelaCPF.Point(algarismo + LeftRequerimento + 2, TopRequerimento + 4) = 0 Then
      soma = soma + 8
    End If
    If pctCopiaPartedaTelaCPF.Point(algarismo + LeftRequerimento + 2, TopRequerimento + 8) = 0 Then
      soma = soma + 16
    End If
    If pctCopiaPartedaTelaCPF.Point(algarismo + LeftRequerimento + 4, TopRequerimento + 2) = 0 Then
      soma = soma + 32
    End If
    If pctCopiaPartedaTelaCPF.Point(algarismo + LeftRequerimento + 4, TopRequerimento + 6) = 0 Then
      soma = soma + 64
    End If
    Select Case soma
      Case 28
        letra = "1"
      Case 52
        letra = "2"
      Case 124
        letra = "3"
      Case 66
        letra = "4"
      Case 85
        letra = "5"
      Case 95
        letra = "6"
      Case 12
        letra = "7"
      Case 127
        letra = "8"
      Case 125
        letra = "9"
      Case 119
        letra = "0"
      Case 90
        letra = "INICIAL"
      Case 44
        letra = "PP"
      Case 108
        letra = "PR"
      Case 75
        letra = "NORMAL"
      Case 23
        letra = "DEFERIDO"
      Case 3
        letra = "INDEFERIDO"
      Case 15
        letra = "PENDENTE"
      Case Else
        letra = ""
    End Select
    Requerimento = Requerimento & letra
    algarismo = algarismo + 6
  Next digito
  CapturaNumeroDetalhes = Requerimento
End Function
    
Sub atualizaprogresso()
  On Error Resume Next
  pctProgressoFundo.Width = 2000
  pctProgresso.Width = pctProgressoFundo.Width * GlobalIDRequerimento / GlobalQuantidadedeRequerimentos
  pctProgressoFundo.Top = lblRelogio.Top
  pctProgressoFundo.Left = lblRelogio.Left + lblRelogio.Width + 40
  pctProgressoFundo.Height = lblRelogio.Height
  pctProgressoFundo.Visible = True
  pctProgresso.Cls
  pctProgresso.PSet (0, 0)
  pctProgresso.Top = -20
  pctProgresso.Left = -20
  pctProgresso.Print Space(10) & Format(GlobalIDRequerimento, "00") & " de " & Format(GlobalQuantidadedeRequerimentos, "00")
  pctProgressoFundo.Cls
  pctProgressoFundo.PSet (0, 0)
  pctProgressoFundo.Print Space(10) & Format(GlobalIDRequerimento, "00") & " de " & Format(GlobalQuantidadedeRequerimentos, "00")
  lbVersao.Top = -1000
  lbVersao.Left = Me.Width - lbVersao.Width - 360
End Sub

Sub Escreve(pontox As Long, pontoy As Long, NUMERO As String)
  Dim conta, linha, coluna, digito As Long
  For conta = 1 To Len(NUMERO)
    digito = Mid(NUMERO, conta, 1)
    If digito = 0 Then
      digito = 9
    Else
      digito = digito - 1
    End If
    Picture1.Visible = False
    For linha = 4 To 12
      For coluna = 0 To 6 '160
        If pctDigitos.Point(coluna + 6 * digito + 3, linha) = 2631720 Then
          Picture1.PSet (6 * (conta - 1) + pontox + coluna, pontoy + linha), RGB(0, 0, 0)
        End If
      Next coluna
    Next linha
    Picture1.Visible = True
  Next conta
End Sub

Sub RequerimentonãoEncontrado(Requerimento As String, sequencia As String)
  Dim conta, linha, coluna, digito, pontoy As Long
  
  pontoy = GlobalLinhaPicture
  Picture1.Visible = False
  pctFundoCopias.Top = 3000
  pctFundoCopias.Visible = True
  Picture1.Height = GlobalLinhaPicture * 15 + 300
  If Picture1.Height + 200 > pctFundoCopias.Height Then
    Picture1.Top = 3000
  End If
  Picture1.Width = pctCopiaPartedaTela.Width

  'colore linha
  For linha = 2 To 15
    For coluna = 1 To 726
      Picture1.PSet (coluna, pontoy + linha), RGB(255, 220, 220)
    Next coluna
  Next linha
    
  'escreve o numero do requerimento
  Escreve 2, pontoy, Requerimento
    
  'escreve Este requerimento não foi encontrado
  For linha = 4 To 15
    For coluna = 0 To 178
      If pctEsteRequerimento.Point(coluna, linha) = 0 Then
        Picture1.PSet (80 + coluna, pontoy + linha), RGB(40, 40, 40)
      End If
    Next coluna
  Next linha

  'traça linha preta
  For coluna = 1 To 726
    Picture1.PSet (coluna, pontoy + 16), RGB(40, 40, 40)
    Picture1.PSet (coluna, pontoy + 17), RGB(40, 40, 40)
  Next coluna
  For linha = 2 To 15
    Picture1.PSet (65, pontoy + linha), RGB(40, 40, 40)
    Picture1.PSet (283, pontoy + linha), RGB(40, 40, 40)
    Picture1.PSet (357, pontoy + linha), RGB(40, 40, 40)
    Picture1.PSet (383, pontoy + linha), RGB(40, 40, 40)
    Picture1.PSet (450, pontoy + linha), RGB(40, 40, 40)
    Picture1.PSet (524, pontoy + linha), RGB(40, 40, 40)
    Picture1.PSet (700, pontoy + linha), RGB(40, 40, 40)
    Picture1.PSet (722, pontoy + linha), RGB(40, 40, 40)
    Picture1.PSet (723, pontoy + linha), RGB(40, 40, 40)
    Picture1.PSet (724, pontoy + linha), RGB(40, 40, 40)
    Picture1.PSet (725, pontoy + linha), RGB(40, 40, 40)
    Picture1.PSet (726, pontoy + linha), RGB(40, 40, 40)
  Next linha
  Escreve 703, GlobalLinhaPicture, sequencia
  GlobalLinhaPicture = GlobalLinhaPicture + 16
  SavePicture Picture1.Image, GlobalPastadeTrabalho & "\" & GlobalDatadosRequerimentos & GlobalAgenciaEscolhida & "Todos.bmp"
  Picture1.Visible = True
End Sub

Sub efeitos(Imprime As Boolean, Ordem As String)
  Dim linha, coluna As Long
  Dim sequencia As String
  Dim pontoy As Long

  sequencia = Ordem
  Picture1.Visible = False
  pctFundoCopias.Top = 3000
  pctFundoCopias.Visible = True
  Picture1.Height = GlobalLinhaPicture * 15 + 300
  If Picture1.Height + 200 > pctFundoCopias.Height Then
    Picture1.Top = 3000
  End If
  Picture1.Width = pctCopiaPartedaTela.Width
  If Imprime Then
    For linha = 1 To pctCopiaPartedaTela.Height / 15
      For coluna = 1 To pctCopiaPartedaTela.Width / 15
        If pctCopiaPartedaTela.Point(coluna, linha) = 0 Then
          Picture1.PSet (coluna, GlobalLinhaPicture + linha + 1), RGB(220, 220, 220)
        Else
          Picture1.PSet (coluna, GlobalLinhaPicture + linha + 1), RGB(40, 40, 40)
        End If
      Next coluna
    Next linha
  Else
    For linha = 1 To pctCopiaPartedaTela.Height / 15 - 2
      For coluna = 1 To pctCopiaPartedaTela.Width / 15
        If pctCopiaPartedaTela.Point(coluna, linha) = 0 Then
          Picture1.PSet (coluna, GlobalLinhaPicture + linha + 1), RGB(255, 255, 225)
        Else
          Picture1.PSet (coluna, GlobalLinhaPicture + linha + 1), RGB(120, 120, 120)
        End If
      Next coluna
    Next linha
    For coluna = 1 To pctCopiaPartedaTela.Width / 15
      Picture1.PSet (coluna, GlobalLinhaPicture + linha + 1), RGB(40, 40, 40)
      Picture1.PSet (coluna, GlobalLinhaPicture + linha + 2), RGB(40, 40, 40)
    Next coluna
  End If
  pontoy = GlobalLinhaPicture
  Escreve 703, pontoy, sequencia
  GlobalLinhaPicture = GlobalLinhaPicture + 16
  SavePicture Picture1.Image, GlobalPastadeTrabalho & "\" & GlobalDatadosRequerimentos & GlobalAgenciaEscolhida & "Todos.bmp"
  Picture1.Visible = True
End Sub
  
Private Function RequerimentosAgendaAnterior(nomeArquivo As String) As Long
  Dim FileNumber  As Long
  Dim mTexto, mLinha As String
  Dim contador As Long
  Dim ultimoNIT As Long
  Dim pos1, pos2, pos3, pos4, pos5, pos6, pos7, pos8 As Long
    
  contador = 0
  FileNumber = FreeFile
  Open GlobalPastadeTrabalho & "\" & nomeArquivo For Input As #FileNumber
  Do While Not EOF(FileNumber)
    Line Input #FileNumber, mLinha
    If Len(mLinha) = 0 Then Exit Do
  contador = contador + 1
  GlobalQuantidadedeRequerimentos = contador
  pos1 = InStr(1, mLinha, Chr(9))
    If pos1 > 0 Then
    GlobalRequerimentos(contador).sequencia = Mid(mLinha, 1, pos1 - 1)
    pos2 = InStr(pos1 + 1, mLinha, Chr(9))
    If pos2 > 0 Then
      GlobalRequerimentos(contador).Número = Mid(mLinha, pos1 + 1, pos2 - pos1 - 1)
    End If
    pos3 = InStr(pos2 + 1, mLinha, Chr(9))
    If pos3 > 0 Then
      GlobalRequerimentos(contador).Tipo = Mid(mLinha, pos2 + 1, pos3 - pos2 - 1)
      pos4 = InStr(pos3 + 1, mLinha, Chr(9))
      If pos4 > 0 Then
        GlobalRequerimentos(contador).Status = Mid(mLinha, pos3 + 1, pos4 - pos3 - 1)
        pos5 = InStr(pos4 + 1, mLinha, Chr(9))
        If pos5 > 0 Then
          GlobalRequerimentos(contador).nit = Mid(mLinha, pos4 + 1, pos5 - pos4 - 1)
          If GlobalRequerimentos(contador).nit <> "" Then ultimoNIT = contador
            pos6 = InStr(pos5 + 1, mLinha, Chr(9))
            If pos6 > 0 Then
              GlobalRequerimentos(contador).Impresso = Mid(mLinha, pos5 + 1, pos6 - pos5 - 1)
              pos7 = InStr(pos6 + 1, mLinha, Chr(9))
              If pos7 > 0 Then
                GlobalRequerimentos(contador).Segurado = Mid(mLinha, pos6 + 1, pos7 - pos6 - 1)
                GlobalRequerimentos(contador).Crítica = Mid(mLinha, pos7 + 1)
              End If
            End If
          End If
        End If
      End If
    End If
    mTexto = mTexto & mLinha & Chr(13) & Chr(10)
  Loop
  Close #FileNumber
  RequerimentosAgendaAnterior = ultimoNIT
End Function

Sub ColocaTelaControleOperacionanoModoNormal()
  On Error Resume Next
  
  If GlobalIDControleOperacional Then
    If AppToForeground(, GlobalIDControleOperacional, SW_NORMAL) Then
        ' a tela foi maximizada
    Else
      MsgBox "Falha ao maximizar a tela Controle Operacional", vbCritical, "Maximizar Controle Operacional"
    End If
  Else
    MsgBox "Não foi encontrada tela Controle Operaciona", vbCritical, "Maximizar Controle Operacional"
  End If
End Sub

Sub ConverteData(datalonga As String)
  On Error Resume Next
  Dim memo As String
  Dim dia As Long
  Dim mes As Long
  Dim ano As Long
  
  GlobalDatadosRequerimentos = "00000000"
  memo = UCase(datalonga)
  If InStr(1, memo, ",") Then memo = Trim(Mid(memo, InStr(1, memo, ",") + 1))
  dia = Val(memo)
  ano = Val(Right$(memo, 4))
  If InStr(1, memo, "DE") Then memo = Trim(Mid(memo, InStr(1, memo, "DE") + 3))
  If InStr(1, memo, "DE") Then memo = Trim(Mid(memo, 1, InStr(1, memo, "DE") - 1))
  Select Case memo
    Case "JANEIRO"
      mes = 1
    Case "FEVEREIRO"
      mes = 2
    Case "MARÇO"
      mes = 3
    Case "ABRIL"
      mes = 4
    Case "MAIO"
      mes = 5
    Case "JUNHO"
      mes = 6
    Case "JULHO"
      mes = 7
    Case "AGOSTO"
      mes = 8
    Case "SETEMBRO"
      mes = 9
    Case "OUTUBRO"
      mes = 10
    Case "NOVEMBRO"
      mes = 11
    Case "DEZEMBRO"
      mes = 12
  End Select
  GlobalDatadosRequerimentos = Format(ano, "0000") & Format(mes, "00") & Format(dia, "00")
End Sub

Function DialogGetHwnd(Optional ByVal sDialogCaption As String = vbNullString, Optional sClassName As String = vbNullString) As Long
  On Error Resume Next
  
  DialogGetHwnd = FindWindowA(sClassName, sDialogCaption)
  On Error GoTo 0
End Function

Sub atualizarStatus()
  Me.Left = 600
  If GlobalAgenciaEscolhida = "" Then
    lbNomePrograma.Caption = "Agendamentos do dia " & Mid(GlobalDatadosRequerimentos, 7, 2) & "/" & Mid(GlobalDatadosRequerimentos, 5, 2) & "/" & Mid(GlobalDatadosRequerimentos, 1, 4)
  Else
    lbNomePrograma.Caption = Mid(GlobalAgenciaEscolhida, 1, 40) & ", " & Mid(GlobalDatadosRequerimentos, 7, 2) & "/" & Mid(GlobalDatadosRequerimentos, 5, 2) & "/" & Mid(GlobalDatadosRequerimentos, 1, 4)
  End If
End Sub

Private Sub ClickOpen(hMsgBox As Long)
  On Error Resume Next
  Dim hButtonOpen As Long
  Dim hComboBox As Long
  
  hButtonOpen = FindWindowEx(hMsgBox, 0, "Button", "OK")
  SendMessage hButtonOpen, BM_CLICK, 0, 0
End Sub

Private Function Crystal(hMsgBox As Long) As Boolean
  On Error Resume Next
  Dim h1AfxWnd42 As Long
  Dim h2AfxWnd42 As Long
  Dim hAfxFrameOrView42 As Long
  
  h1AfxWnd42 = FindWindowEx(hMsgBox, 0, "AfxWnd42", "")
  h2AfxWnd42 = FindWindowEx(h1AfxWnd42, 0, "AfxWnd42", "")
  hAfxFrameOrView42 = FindWindowEx(h2AfxWnd42, 0, "AfxFrameOrView42", "")
  Crystal = hAfxFrameOrView42 <> 0
End Function

Sub MontaListadeRequerimentos(memotexto As String)
  Dim pos As Long
  Dim linha As String
  Dim conta As Long
  Dim indice As Long
  Dim nomedosegurado As String
  Dim GlobalRequerimentosProv(1000) As Requerimento
  Dim LINHA2 As String
  Dim pos3 As Long
  Dim conta2 As Long
  
  indice = 0
  conta = 0
  GlobalQuantidadedeRequerimentos = GlobalAgendamentosQuandidade
  lstClassificar.Clear
  If optOrdem(1).Value = True Then
    For conta = 1 To GlobalAgendamentosQuandidade
      lstClassificar.AddItem Mid(GlobalAgendamentosConsulta(conta).Horario, 1, 5) & " - " & GlobalAgendamentosConsulta(conta).Segurado
      lstClassificar.ItemData(lstClassificar.NewIndex) = GlobalAgendamentosConsulta(conta).Requerimento
    Next conta
  Else
    For conta = 1 To GlobalAgendamentosQuandidade
      lstClassificar.AddItem GlobalAgendamentosConsulta(conta).Segurado
      lstClassificar.ItemData(lstClassificar.NewIndex) = GlobalAgendamentosConsulta(conta).Requerimento
    Next conta
  End If
  lstMostrarRequerimentos.Clear
  lstMostrarRequerimentos.AddItem "Seq." & Chr(9) & "Requerim." & Chr(9) & Chr(9) & Chr(9) & "Segurado"
  For conta = 0 To lstClassificar.ListCount - 1
    For conta2 = 1 To GlobalAgendamentosQuandidade
      If GlobalAgendamentosConsulta(conta2).Requerimento = (lstClassificar.ItemData(conta)) Then
        Exit For
      End If
    Next conta2
    indice = indice + 1
    GlobalRequerimentos(conta + 1).Número = lstClassificar.ItemData(conta)
    GlobalRequerimentos(conta + 1).Segurado = lstClassificar.List(conta)
    lstMostrarRequerimentos.AddItem Format(indice, "000") & Chr(9) & lstClassificar.ItemData(conta) & Chr(9) & lstClassificar.List(conta)
    
    'valores iniciais
    GlobalRequerimentos(conta).Tipo = ""
    GlobalRequerimentos(conta).Status = ""
    GlobalRequerimentos(conta).nit = ""
    GlobalRequerimentos(conta).Crítica = ""
  Next conta
  lstMostrarRequerimentos.Visible = True
  If GlobalQuantidadedeRequerimentos > 0 Then
    LocalCopiar = True
    Me.Top = Screen.Height - 760 - 3000
    Me.Width = 12540 + 1600
    Me.Height = Screen.Height - 560
    SetForegroundWindow (Me.hWnd)
    grupoOrdem.Visible = False
    redimensionarForm -4000, 2500
    fraImprime.Visible = True
    mostralista
    If GlobalQuantidadedeRequerimentos > 40 Then
      paracima.Visible = True
      parabaixo.Visible = True
    End If
    editMarcacaoPara.Text = GlobalQuantidadedeRequerimentos
  Else
    res = SetWindowPos(Me.hWnd, -2, 0, 0, 0, 0, 3)
    MsgBox "Não foi encontrado nenhum agendamento de perícia para esta data" & Chr(13) & Chr(10) & GlobalAgendamentosConsultaCabecalho, vbCritical, "Agendamentos do SABI"
    End
  End If
 
End Sub

Sub MouseClique(posx As Long, posy As Long)
  On Error Resume Next
  Dim Size As RECT
  Dim IDTelaAtiva As Long
  Dim pt As POINTAPI
  
  GetCursorPos pt
  IDTelaAtiva = GetForegroundWindow
  res = GetWindowRect(IDTelaAtiva, Size)
  SetCursorPos Size.Left + posx, Size.Top + posy
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
  SetCursorPos pt.x, pt.y
End Sub
    
Private Sub btoConfirmar_Click()
  On Error Resume Next
  Dim res As String

  Timer2.Enabled = False
  paracima.Visible = False
  parabaixo.Visible = False
  btoIniciar.Enabled = False
  cmdContinua.Enabled = False
  editMarcacaoDe.Enabled = False
  editMarcacaoPara.Enabled = False
  chcExameInicial.Enabled = False
  chcPP.Enabled = False
  lbVersao.Visible = False
  
  If testarSabiAberto Then
    prepararAmbiente
  Else
    MsgBox "Abra o módulo Controle Operacional do SABI e faça o login." & vbCrLf & "O Automatizador irá esperar 60 segundos.", vbInformation, NomeAplicacao
    exibirStatus "Inicie e faça login no Controle Operacional do SABI.", 60
  End If

  abc
  If chcExameInicial.Value = 1 And chcPP.Value = 1 Then Imprimeosrequerimentos ("TODOS")
  If chcExameInicial.Value = 1 And chcPP.Value = 0 Then Imprimeosrequerimentos ("INICIAL")
  If chcExameInicial.Value = 0 And chcPP.Value = 1 Then Imprimeosrequerimentos ("PP")
  If chcExameInicial.Value = 0 And chcPP.Value = 0 Then Imprimeosrequerimentos ("NENHUM")
  SendMessageByLong lstMostrarRequerimentos.hWnd, LB_SETHORIZONTALEXTENT, 1200, 0
  SendMessageByLong lstMostrarRequerimentos.hWnd, WM_VSCROLL, SB_BOTTOM, 0
  SetForegroundWindow (GlobalIDControleOperacional)
  
  Sleep 300
  ClickMenu GlobalIDControleOperacional, 0, 6
  End
End Sub

Private Sub btoFechar_Click()
  End
End Sub

Private Sub btoFechar2_Click()
  Sabi.fecharJanela
  exibirRelatorioFinal
  End
End Sub

Private Sub btoFecharErro_Click()
  btoFechar_Click
End Sub

Private Sub btoIniciar_Click()
  listaClassificar.Visible = False
  listaRequerimentos.Visible = False
  lbVersao.Visible = False
  btoIniciar.Enabled = False
  If testarSabiAberto Then
    prepararSABI
  Else
    MsgBox "Abra o módulo Controle Operacional do SABI e faça o login." & vbCrLf & "O Automatizador irá esperar 60 segundos.", vbInformation, NomeAplicacao
    exibirStatus "Inicie e faça login no Controle Operacional do SABI.", 60
  End If
End Sub

Private Sub btoIniciar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Me.MousePointer = 0
End Sub

Private Sub chcExameInicial_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error Resume Next
  If chcExameInicial.Value = 1 And chcPP.Value = 0 Then SaveSetting "AGENDAMENTODOSABI", "IMPRIMIR", "EXAMES", "INICIAIS"
  If chcExameInicial.Value = 1 And chcPP.Value = 1 Then SaveSetting "AGENDAMENTODOSABI", "IMPRIMIR", "EXAMES", "TODOS"
  If chcExameInicial.Value = 0 And chcPP.Value = 0 Then SaveSetting "AGENDAMENTODOSABI", "IMPRIMIR", "EXAMES", "NENHUM"
  If chcExameInicial.Value = 0 And chcPP.Value = 1 Then SaveSetting "AGENDAMENTODOSABI", "IMPRIMIR", "EXAMES", "PP"
End Sub

Private Sub chcPP_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error Resume Next
  If chcExameInicial.Value = 1 And chcPP.Value = 0 Then SaveSetting "AGENDAMENTODOSABI", "IMPRIMIR", "EXAMES", "INICIAIS"
  If chcExameInicial.Value = 1 And chcPP.Value = 1 Then SaveSetting "AGENDAMENTODOSABI", "IMPRIMIR", "EXAMES", "TODOS"
  If chcExameInicial.Value = 0 And chcPP.Value = 0 Then SaveSetting "AGENDAMENTODOSABI", "IMPRIMIR", "EXAMES", "NENHUM"
  If chcExameInicial.Value = 0 And chcPP.Value = 1 Then SaveSetting "AGENDAMENTODOSABI", "IMPRIMIR", "EXAMES", "PP"
End Sub

Private Sub cmdFechar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Me.MousePointer = 0
End Sub

Private Sub editMarcacaoDe_Change()
  editMarcacaoDe.Text = Val(editMarcacaoDe.Text)
  If Val(editMarcacaoDe.Text) < 1 Then editMarcacaoDe.Text = 1
  If Val(editMarcacaoDe.Text) > GlobalQuantidadedeRequerimentos Then editMarcacaoDe.Text = GlobalQuantidadedeRequerimentos
  btoConfirmar.Enabled = Val(editMarcacaoPara.Text) >= Val(editMarcacaoDe.Text)
End Sub

Private Sub editMarcacaoPara_Change()
  editMarcacaoPara.Text = Val(editMarcacaoPara.Text)
  If Val(editMarcacaoPara.Text) < 1 Then editMarcacaoPara.Text = 1
  If Val(editMarcacaoPara.Text) > GlobalQuantidadedeRequerimentos Then editMarcacaoPara.Text = GlobalQuantidadedeRequerimentos
  btoConfirmar.Enabled = Val(editMarcacaoPara.Text) >= Val(editMarcacaoDe.Text)
End Sub

'Função que é executada quando o form recebe o foco do usuário
Private Sub Form_Activate()
  Dim conta As Long
    
  Picture1.Width = 10890
  For conta = 0 To 10891
    Picture1.PSet (conta, 0), RGB(40, 40, 40)
    Picture1.PSet (conta, 1), RGB(40, 40, 40)
  Next conta
  
  'devido a diferenças da altura da barra de título da janela entre o tema clássico e Windows 7,
  'bloqueia a execução quando o tema clássico estiver ativo.
  If estaTemaAtivo = False Then
    MsgBox "O Automatizador do SABI não suporta o tema clássico do Windows. Personalize a tela do seu computador com o tema 'Windows 7' e execute o Automatizador novamente.", vbCritical, "Tema Aero"
    End
  End If
End Sub

'Função executada quando o form é carregado para memória
Private Sub Form_Load()
  'apenas uma execução por vez
  If App.PrevInstance Then
    MsgBox "O Automatizador do SABI já está em execução. Não é permitido executá-lo duas vezes ao mesmo tempo.", vbCritical, "Agendamentos do SABI"
    End
  End If
        
  'Inicia variáveis globais
  deslocalista = 0
  Picture1.BackColor = RGB(171, 171, 171)
  LocalCopiar = False
  GlobalLinhaPicture = 0
  GlobalModoSimulado = False
  GlobalPrimeiraVez = True
  GlobalRelatorioPronto = False
  pctCopiaPartedaTelaCPF.Top = -1000
  GlobalTítulodaTelaAtiva = ""
  GlobalMenuAtualizado = False
  GlobalIDControleOperacional = 0
  GlobalModoImprimeRequerimentos = False
  GlobalEscalaX = 256 / Screen.Width
  GlobalEscalaX = GlobalEscalaX * 256
  GlobalEscalay = 256 / Screen.Height
  GlobalEscalay = GlobalEscalay * 256
  GlobalInicio = GetTickCount
  GlobalImpressaoAuto = True

  'Consulta os dados salvos no registro e do sistema
  obterDadosRegistro
  obterDadosSistema
    
  'Se a pasta AppData da aplicacao nao existir, crie-a
  If Dir(GlobalPastadeTrabalho, vbDirectory) = "" Then
    MkDir GlobalPastadeTrabalho
  End If
  
  'excluir arquivos da pasta temporária
  excluirArquivosTemp

  'Altura e Posicao superior da janela
  redimensionarForm -3000, 2000
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim lngReturnValue As Long
  
  Me.MousePointer = 9
  Call ReleaseCapture
  lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub Form_Resize()
  If LocalCopiar Then
    Me.Width = 8145
    Me.Left = 600
    lbVersao.Top = 0
  Else
    lbVersao.Top = Me.Height - 1590
  End If
  pctFundo.Top = 0
  pctFundo.Left = 0
  pctFundo.Width = Me.Width
  pctFundo.Height = Me.Height
  lbNomePrograma.Top = 30
  lbNomePrograma.Left = Me.Width / 2 - lbNomePrograma.Width / 2 + imageIcone.Left + imageIcone.Width
  lbVersao.Left = Me.Width / 2 - lbVersao.Width / 2 + imageIcone.Left + imageIcone.Width
  grupoOrdem.Left = 120
  grupoOrdem.Top = Me.Height - grupoOrdem.Height - 470
  
  lstMostrarRequerimentos.Top = imageIcone.Top + imageIcone.Height + 40
  lstMostrarRequerimentos.Left = 240
  lstMostrarRequerimentos.Width = Me.Width - 580
  lstMostrarRequerimentos.Height = Abs(pctFundo.Height - lstMostrarRequerimentos.Top - 600)
  fraImprime.Top = grupoOrdem.Top
  fraImprime.Left = 120
  pctCopiaPartedaTela.Left = 0
  pctCopiaPartedaTela.Top = pctFundo.Height + 1000
  lblRelogio.Top = lbNomePrograma.Top + 40
  lblRelogio.Left = imageIcone.Left + imageIcone.Width + 40
  Picture1.Left = 0
  pctFundoCopias.Top = 3000
  pctFundoCopias.Left = 0
  pctFundoCopias.Width = pctFundo.Width
  pctFundoCopias.Height = Abs(pctFundo.Height - 1350)
  Picture1.Left = 0
  lblLocaleData.Left = Me.Width / 2 - 2000
  lblLocaleData.Width = Me.Width / 2 + 2000
  lblLocaleData.Top = 80
  lblLocaleData.AutoSize = True
End Sub

Private Sub lblImprimir_Click()
  Dim conta As Long
  Dim Size As RECT
  
  SetForegroundWindow (GlobalIDControleOperacional)
  For conta = 1 To GlobalQuantidadedeRequerimentos
    If GlobalRequerimentos(conta).nit <> "" Then
      'Abre tela Segunda Via de Marcação de Exame
      SetForegroundWindow (GlobalIDControleOperacional)
      ClickMenu GlobalIDControleOperacional, 4, 7
      espera 1000
      SimulaSendKeys "<TAB>"
      espera 100
      SimulaSendKeys "<TAB>"
      espera 100
      SimulaSendKeys GlobalRequerimentos(conta).nit
      espera 100
      SimulaSendKeys Left$(GlobalRequerimentos(conta).nit, 1)
      espera 300
      MsgBox "imprime"
    End If
  Next conta
End Sub

Private Sub lblImprimir_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  lblImprimir.BorderStyle = 1
End Sub

Private Sub lblImprimir_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  lblImprimir.BorderStyle = 0
End Sub

Private Sub lblRelogio_Change()
  lblRelogio.Visible = lblRelogio.Caption <> "00"
  lblRelogio.Left = imageIcone.Left + imageIcone.Width + 40
  pctProgressoFundo.Left = lblRelogio.Left + lblRelogio.Width + 40
End Sub

Private Sub lblRelogio_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim lngReturnValue As Long
  
  Me.MousePointer = 9
  Call ReleaseCapture
  lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  Me.Top = Screen.Height - 760 - 3000
End Sub

Private Sub lblversao_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim lngReturnValue As Long
  
  Me.MousePointer = 9
  Call ReleaseCapture
  lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  Me.Top = Screen.Height - 760 - 3000
End Sub

Private Sub lblversao_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Me.MousePointer = 0
End Sub

Private Sub lstMostrarRequerimentos_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim lngReturnValue As Long
  
  imprimereq = True
  Me.MousePointer = 9
  Call ReleaseCapture
  lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  Me.Top = Screen.Height - 760 - 3000
End Sub

Private Sub lstMostrarRequerimentos_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error Resume Next
  Dim memo As String
  Dim memorequerimento As String
  Dim numerorequerimento As Long
  Dim pos As Long
  
  If GlobalRelatorioPronto Then
    Me.MousePointer = 0
    lstMostrarRequerimentos.ListIndex = ItemUnderMouse(lstMostrarRequerimentos.hWnd, x, y)
    pos = ItemUnderMouse(lstMostrarRequerimentos.hWnd, x, y)
    If pos > 0 Then
      memorequerimento = lstMostrarRequerimentos.List(pos)
      memorequerimento = Trim(Mid(memorequerimento, 4))
      pos = InStr(1, memorequerimento, " ")
      If pos > 0 Then
        memorequerimento = Mid(memorequerimento, 1, pos - 1)
        If GlobalRequerimentoMostrado <> memorequerimento Then
          GlobalRequerimentoMostrado = memorequerimento
          memo = Dir(GlobalPastadeTrabalho & "\*" & memorequerimento & ".bmp")
          If memo <> "" Then
            pctApresentaPartedaTelaCopiada.Picture = LoadPicture(GlobalPastadeTrabalho & "\" & memo)
            pctApresentaPartedaTelaCopiada.Top = 0
            pctApresentaPartedaTelaCopiada.Left = 0
            pctApresentaPartedaTelaCopiada.Visible = True
            pctApresentaPartedaTelaCopiada.ZOrder
            Me.Top = Screen.Height - 760 - 3000
            DoEvents
          Else
            pctApresentaPartedaTelaCopiada.Visible = False
          End If
        End If
      End If
    End If
  End If
End Sub

Private Sub lstMostrarRequerimentos_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  imprimereq = False
End Sub

Private Sub MontaRelaçãodeRequerimentos()
  On Error Resume Next
  Dim COsize As RECT
  Dim Size As RECT
  Dim titletmp As String
  Dim nret As Long
  Dim TelaSize As RECT
  Dim arquivo As String
  Dim memo As String
  Dim contador As Integer
  
  On Error Resume Next

  GlobalSeçao = Format(Date, "YYYYMMDD") & Format(Time, "hhmmss")
  GlobalNomedoRelatorio = "requerimentos" & GlobalSeçao
  
  'espera a proxima tela (a tela do crystal report não tem nome)
  While GlobalIDTelaImprimirAgendamento = GetForegroundWindow
    espera 100
    DoEvents
  Wend
  arquivo = esperaCRYSTALREPORTeExporta
End Sub

Private Sub optOrdem_Click(Index As Integer)
  If Index = 1 Then
    SaveSetting "AGENDAMENTODOSABI", "IMPRIMIR", "ORDEM", "HORA"
  Else
    SaveSetting "AGENDAMENTODOSABI", "IMPRIMIR", "ORDEM", "NOME"
  End If
End Sub

Private Sub parabaixo_Click()
  Dim conta As Long

  mtempo2 = 0
  deslocalista = deslocalista - 5
  If deslocalista < 0 Then deslocalista = 0
  lstMostrarRequerimentos.Clear
  For conta = 1 To 50
    If (conta + deslocalista) < GlobalQuantidadedeRequerimentos + 1 Then
      lstMostrarRequerimentos.AddItem Format(conta + deslocalista, "000") & "         " & GlobalRequerimentos(conta + deslocalista).Número & "         " & GlobalRequerimentos(conta + deslocalista).Segurado
    End If
  Next conta
  lstMostrarRequerimentos.Height = Screen.Height
  mostralista
End Sub

Private Sub paracima_Click()
  Dim conta As Long

  mtempo2 = 0
  deslocalista = deslocalista + 5
  If deslocalista > GlobalQuantidadedeRequerimentos Then deslocalista = GlobalQuantidadedeRequerimentos
  lstMostrarRequerimentos.Visible = False
  lstMostrarRequerimentos.Clear
  For conta = 1 To 50
    If (conta + deslocalista) < GlobalQuantidadedeRequerimentos + 1 Then
      lstMostrarRequerimentos.AddItem Format(conta + deslocalista, "000") & "         " & GlobalRequerimentos(conta + deslocalista).Número & "         " & GlobalRequerimentos(conta + deslocalista).Segurado
    End If
  Next conta
  lstMostrarRequerimentos.Height = Screen.Height
  mostralista
End Sub

Private Sub pctCopiaPartedaTela_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim lngReturnValue As Long

  Me.MousePointer = 9
  Call ReleaseCapture
  lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  If Me.Height = 133 * 15 Or LocalCopiar = True Then Me.Top = Screen.Height - 760 - 3000
End Sub

Private Sub pctFundo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim lngReturnValue As Long
  
  Me.MousePointer = 9
  Call ReleaseCapture
  lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  If Me.Height = 133 * 15 Or LocalCopiar = True Then Me.Top = Screen.Height - 760 - 3000
End Sub

Private Sub pctFundo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Me.MousePointer = 0
End Sub

Private Sub pctCopiaPartedaTela_Change()
  GlobalTipo = "PP/PR"
  GlobalAlerta = False
  pctCopiaPartedaTela.Top = 0
  DoEvents
  If Point(pctCopiaPartedaTela.Left + 790 * 15, pctCopiaPartedaTela.Top + 7 * 15 - 45) = 16777215 Then GlobalAlerta = True
  If Point(pctCopiaPartedaTela.Left + 166 * 15 + 8, pctCopiaPartedaTela.Top + 53 * 15 - 45) <> 0 Then GlobalTipo = "INICIAL"
  pctCopiaPartedaTela.Top = pctFundo.Height + 1000
End Sub

Private Sub pctCopiaPartedaTela_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim lngReturnValue As Long

  Me.MousePointer = 5
  Call ReleaseCapture
  lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  Me.Top = Screen.Height - 760 - 3000
End Sub

Private Sub pctFundoCopias_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim lngReturnValue As Long
  
  Me.MousePointer = 9
  Call ReleaseCapture
  lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  Me.Top = Screen.Height - 760 - 3000
End Sub

Private Sub pctFundoCopias_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Me.MousePointer = 0
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error Resume Next
  
  Clipboard.Clear
  Clipboard.SetData Picture1.Picture
  LocalY = y
  If Button = 2 Then
    res = SetWindowPos(Me.hWnd, -2, 0, 0, 0, 0, 3)
    MsgBox "Esta relação de requerimentos foi salva na área de transferência. Para imprimir cole agora no 'Word' ou 'Paint'.", vbCritical, "Imprimir Relação de Requerimentos"
    res = SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
  End If
End Sub

Private Sub timerAbrirSabi_Timer()
  On Error Resume Next
  
  txtStatusAguarda.Caption = "Aguardando " & contadorTimer & " segundos..."
  contadorTimer = contadorTimer - 1
  If contadorTimer < 0 Then
    If testarSabiAberto Then
      pctFundo.Visible = True
      painelStatus.Visible = False
      prepararSABI
    Else
      MsgBox "O Automatizador não conseguiu encontrar o Controle Operacional do SABI aberto. Se o SABI estiver apresentando lentidão, tente novamente mais tarde.", vbCritical, NomeAplicacao
      End
    End If
  End If
End Sub

Private Sub Timer2_Timer()
  On Error Resume Next
  
  mtempo2 = mtempo2 + 1
  If mtempo2 > 60 Then End
End Sub

Private Sub txtUltimo_Change()
  txtUltimo.Text = Val(txtUltimo.Text)
  If Val(txtUltimo.Text) < 1 Then txtUltimo.Text = 1
  If Val(txtUltimo.Text) > GlobalQuantidadedeRequerimentos Then txtUltimo.Text = GlobalQuantidadedeRequerimentos
  cmdContinua.Enabled = Val(txtUltimo.Text) >= Val(editMarcacaoDe.Text)
End Sub
