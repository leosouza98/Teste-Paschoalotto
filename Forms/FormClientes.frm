VERSION 5.00
Begin VB.Form FormClientes 
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   7320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   10755
   StartUpPosition =   2  'CenterScreen
   Tag             =   "+"
   Begin VB.Frame Consulta 
      Height          =   945
      Left            =   90
      TabIndex        =   9
      Top             =   60
      Width           =   10605
      Begin VB.Frame FrameCodigo 
         Caption         =   "Codigo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1020
         TabIndex        =   19
         Top             =   150
         Width           =   1365
         Begin VB.TextBox Codigo 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   90
            TabIndex        =   2
            Top             =   210
            Width           =   1155
         End
      End
      Begin VB.CommandButton LimpaDados 
         Caption         =   "Limpa"
         Height          =   585
         Left            =   5790
         Picture         =   "FormClientes.frx":0000
         TabIndex        =   4
         Top             =   210
         Width           =   735
      End
      Begin VB.CommandButton bt_Salva 
         Caption         =   "Salva"
         Height          =   585
         Left            =   4620
         Picture         =   "FormClientes.frx":23E6
         TabIndex        =   3
         Top             =   210
         Width           =   735
      End
      Begin VB.CommandButton bt_novo 
         Caption         =   "Novo"
         Height          =   585
         Left            =   120
         Picture         =   "FormClientes.frx":47CC
         TabIndex        =   1
         Top             =   210
         Width           =   735
      End
   End
   Begin VB.Frame FrameDados 
      Caption         =   "Dados do Cliente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   90
      TabIndex        =   0
      Top             =   1050
      Width           =   10605
      Begin VB.Frame Frame 
         Caption         =   "Importar Clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   3720
         TabIndex        =   21
         Top             =   3360
         Width           =   4215
         Begin VB.TextBox caminhoarquivo 
            Enabled         =   0   'False
            Height          =   375
            Left            =   150
            TabIndex        =   24
            Top             =   360
            Width           =   2565
         End
         Begin VB.CommandButton bt_importa 
            Caption         =   "Importar"
            Height          =   585
            Left            =   2790
            Picture         =   "FormClientes.frx":6ACC
            TabIndex        =   23
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton localizararquivo 
            Caption         =   "..."
            Height          =   315
            Left            =   3600
            TabIndex        =   22
            Top             =   360
            Width           =   315
         End
      End
      Begin VB.Frame Filtro 
         Caption         =   "Exportar Clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   150
         TabIndex        =   18
         Top             =   3360
         Width           =   1905
         Begin VB.CommandButton bt_exporta 
            Caption         =   "Exportar"
            Height          =   585
            Left            =   450
            Picture         =   "FormClientes.frx":8EB2
            TabIndex        =   20
            Top             =   270
            Width           =   735
         End
      End
      Begin VB.Frame FramObs 
         Caption         =   "Observação:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   150
         TabIndex        =   17
         Top             =   2010
         Width           =   10335
         Begin VB.TextBox Observacao 
            Height          =   945
            Left            =   90
            MaxLength       =   100
            TabIndex        =   8
            Top             =   240
            Width           =   10125
         End
      End
      Begin VB.Frame FrameDataNascimento 
         Caption         =   "Data Nascimento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   150
         TabIndex        =   14
         Top             =   1140
         Width           =   1815
         Begin VB.TextBox DataNascimento 
            Alignment       =   2  'Center
            Height          =   405
            Left            =   90
            MaxLength       =   11
            TabIndex        =   7
            Top             =   240
            Width           =   1635
         End
         Begin VB.Frame Frame4 
            Caption         =   "Nome:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   -30
            TabIndex        =   15
            Top             =   840
            Width           =   5565
            Begin VB.TextBox Text3 
               Height          =   405
               Left            =   90
               TabIndex        =   16
               Top             =   240
               Width           =   5295
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "CPF/CNPJ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   5820
         TabIndex        =   11
         Top             =   330
         Width           =   3735
         Begin VB.TextBox CPF_CNPJ 
            Height          =   405
            Left            =   120
            MaxLength       =   14
            TabIndex        =   6
            Top             =   240
            Width           =   3495
         End
      End
      Begin VB.Frame FrameNomeCliente 
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   150
         TabIndex        =   10
         Top             =   330
         Width           =   5565
         Begin VB.Frame Frame2 
            Caption         =   "Nome:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   -30
            TabIndex        =   12
            Top             =   840
            Width           =   5565
            Begin VB.TextBox Text2 
               Height          =   405
               Left            =   90
               TabIndex        =   13
               Top             =   240
               Width           =   5295
            End
         End
         Begin VB.TextBox NomeCliente 
            Height          =   405
            Left            =   90
            MaxLength       =   50
            TabIndex        =   5
            Top             =   240
            Width           =   5295
         End
      End
   End
End
Attribute VB_Name = "FormClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim glCodigoCliente As String, gldataformat As Date, glarquivo As String
Dim conexao As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim cursor As New ADODB.Recordset
Private Sub Form_Load()

   conexao.ConnectionString = "Driver={MySQL ODBC 3.51 Driver};Server=Localhost;Port=3306;Database=paschoalloto;User=root;Password=123panoramix"
   conexao.CursorLocation = adUseClient
   conexao.Open
   Set cmd.ActiveConnection = conexao
                      
End Sub

Private Sub bt_novo_Click()
    Dim comando As String, vCodigoCliente As String

    comando = "Select ifnull(Max(codigoCliente),0) as max_codigo from clientes;"
    cursor.Open comando, conexao
    
    If Not cursor.EOF Then
        vCodigoCliente = cursor.Fields("max_codigo").Value + 1
    End If
    
    cursor.Close

    If vCodigoCliente = "" Then
        vCodigoCliente = 1
        glCodigoCliente = vCodigoCliente
    Else
        glCodigoCliente = vCodigoCliente
    End If
    
    Codigo.Text = glCodigoCliente
    
End Sub

Private Sub bt_Salva_Click()
    Dim comando As String
    
    gldataformat = Date
    DataNascimento = Format(gldataformat, "yyyy-mm-dd")
    
    If validacadatro = True Then
        
        comando = "Select CodigoCliente from clientes where codigocliente = " & Codigo & ";"
        cursor.Open comando, conexao
        If Not cursor.EOF Then
            comando = "UPDATE clientes SET nome = '" & NomeCliente & "', CPF_CNPJ = '" & CPF_CNPJ & "',DataNascimento = '" & DataNascimento & "', Observacao ='" & Observacao & "' Where CodigoCliente = " & Codigo & ""
            conexao.Execute comando
            MsgBox "Dados atualizados com Sucesso!"
        Else
            comando = "Insert into clientes (CodigoCliente,Nome,CPF_CNPJ,DataNascimento,Observacao) values (" & glCodigoCliente & " , '" & NomeCliente & "', '" & CPF_CNPJ & "','" & DataNascimento & "', '" & Observacao & "')"
            conexao.Execute comando
            MsgBox "Dados criados com Sucesso!"
        End If
    End If

    cursor.Close
End Sub
Private Function validacadatro() As Boolean
    validacadatro = False
    
    If NomeCliente = "" Then
        MsgBox "Insira um nome!"
        Exit Function
    End If
    
    If CPF_CNPJ = "" Then
        MsgBox "Insira um CPF/CNPJ!"
        Exit Function
    End If
    
    If DataNascimento = "" Then
        MsgBox "Insira uma Data de Nascimento!"
        Exit Function
    End If
    
    validacadatro = True

End Function

Private Sub Codigo_GotFocus()
    If Not Codigo = "" Then
        BuscaDadosBasicos
    Else
        Exit Sub
    End If
End Sub

Sub BuscaDadosBasicos()
    Dim comando As String
    Dim cursor As New ADODB.Recordset

    comando = "Select CodigoCliente,Nome,CPF_CNPJ,DataNascimento,Observacao from clientes where codigocliente = " & Codigo
    cursor.Open comando, conexao
    
    If Not cursor.EOF Then
        Codigo.Text = cursor.Fields("CodigoCliente").Value
        NomeCliente.Text = cursor("Nome")
        CPF_CNPJ.Text = cursor("CPF_CNPJ")
        DataNascimento.Text = cursor("DataNascimento")
        Observacao.Text = cursor("Observacao")
    End If
    
End Sub

Private Sub DataNascimento_LostFocus()
   Dim data As Date
   
   If IsDate(DataNascimento.Text) Then
      data = CDate(DataNascimento.Text)
   Else
      MsgBox "Insira uma data válida no formato dd/mm/yyyy."
   End If
End Sub

Private Sub LimpaDados_Click()
    Codigo = ""
    NomeCliente = ""
    CPF_CNPJ = ""
    DataNascimento = ""
    Observacao = ""
    Lista = ""
End Sub

Private Sub localizararquivo_Click()
    Dim arquivo As String
    arquivo = BuscarArquivo()
    If arquivo <> "" Then
        caminhoarquivo.Text = arquivo
        glarquivo = arquivo
    End If
End Sub
Private Sub bt_importa_Click()


    If caminhoarquivo = "" Then
        MsgBox "Selecione um aquivo para ser importado!"
        Exit Sub
    End If

    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim i As Integer
    Dim codigocliente As Integer
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(glarquivo)
    Set xlSheet = xlBook.Worksheets("Planilha1")
    
    cursor.Open "SELECT * FROM clientes", conexao, adOpenStatic, adLockOptimistic
    
    For i = 1 To xlSheet.UsedRange.Rows.Count
            cursor.AddNew
            cursor("CodigoCliente") = xlSheet.Cells(i, 1)
            cursor("nome") = xlSheet.Cells(i, 2)
            cursor("CPF_CNPJ") = xlSheet.Cells(i, 3)
            cursor("DataNascimento") = xlSheet.Cells(i, 4)
            cursor("Observacao") = xlSheet.Cells(i, 5)
            cursor.Update
    Next i
    
    MsgBox "Dados importados com sucesso!"
    
    xlBook.Close
    xlApp.Quit
    
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
End Sub
Function BuscarArquivo() As String
    Dim arquivo As String
    Dim dialogo As Object
    Set dialogo = CreateObject("MSComDlg.CommonDialog")
    dialogo.MaxFileSize = 260
    dialogo.DialogTitle = "Selecione um arquivo"
    dialogo.ShowOpen
    If dialogo.FileName <> "" Then
        arquivo = dialogo.FileName
    End If
    Set dialogo = Nothing
    BuscarArquivo = arquivo
End Function

Private Sub NomeCliente_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 65 And KeyAscii <= 90) And _
       Not (KeyAscii >= 97 And KeyAscii <= 122) And _
       Not (KeyAscii >= 32 And KeyAscii <= 47) And _
       Not (KeyAscii >= 58 And KeyAscii <= 64) And _
       Not (KeyAscii >= 91 And KeyAscii <= 96) And _
       Not (KeyAscii >= 123 And KeyAscii <= 126) And _
       Not (KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub CPF_CNPJ_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And _
       Not (KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If Not Codigo = "" Then
        BuscaDadosBasicos
    Else
        Exit Sub
    End If
End Sub

