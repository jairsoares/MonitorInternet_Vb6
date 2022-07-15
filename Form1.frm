VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "CSS-Sistemas"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7800
      Top             =   2280
   End
   Begin VB.Label Label2 
      Caption         =   "Status: Parado"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Verificador de Internet Estável"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim caminho As String

caminho = App.Path & "\logInternet.log"
If Dir(caminho) = "" Then     ' Não tem Arquivo de LOG Criado
   ' Vamos Criar o Arquivo de LOG Inicial
   '--------------------------------------
   'define o objeto filesystem e demais variaveis
   
   Dim fso As New FileSystemObject
      
  

   'se o arquivo não existir então cria
   If Not fso.FileExists(caminho) Then
      
      Dim arquivo As File
      Dim arquivoLog As TextStream
      Dim msg As String
      
      Set arquivoLog = fso.CreateTextFile(caminho)
      arquivoLog.Close
      Set arquivo = fso.GetFile(caminho)
      
      Set arquivoLog = arquivo.OpenAsTextStream(ForAppending)
      msg = "CSS-Sistemas Software House" & vbNewLine & _
            "Arquivo de Log de Monitoria de disponibilidade de Internet." & vbNewLine & _
            "------------------------------------------------------------------------------------------"

      arquivoLog.WriteLine msg
      arquivoLog.Close
      Set arquivoLog = Nothing
      Set fso = Nothing
   End If
End If

End Sub

Private Sub Timer1_Timer()

If VerificaInternet() = 0 Then
   Me.Label2 = "Ultimo Status: Queda no Sinal ( " & Date & " as " & Time & " ) "
   registraLogErros 0, "Queda de Sinal Internet", "Servidor"
Else
   DoEvents
   If pingWS() Then
      Me.Label2 = "Ultimo Status: Internet OK"
   Else
      Me.Label2 = "Ultimo Status: Falha no Ping"
   End If
End If

End Sub
