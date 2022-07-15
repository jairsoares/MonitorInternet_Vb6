Attribute VB_Name = "Log"
Option Explicit
Public caminho As String

Public Sub registraLogErros(ByVal lNumero As Long, ByVal sDescricao As String, ByVal sOrigem As String)

'define o objeto filesystem e demais variaveis
Dim fso As New FileSystemObject
Dim arquivo As File
Dim arquivoLog As TextStream
Dim msg As String
caminho = App.Path & "\logInternet.log"

'se o arquivo não existir então cria
If fso.FileExists(caminho) Then
   Set arquivo = fso.GetFile(caminho)
Else
   Set arquivoLog = fso.CreateTextFile(caminho)
   arquivoLog.Close
   Set arquivo = fso.GetFile(caminho)
End If

'prepara o arquivo para anexa os dados
Set arquivoLog = arquivo.OpenAsTextStream(ForAppending)

'monta informações para gerar a linha com erro
msg = "[" & Now() & "]" & sDescricao

' inclui linhas no arquivo texto
arquivoLog.WriteLine msg
' escreve uma linha em branco no arquivo - se voce quiser
'arquivoLog.WriteBlankLines (1)

'fecha e libera o objeto
arquivoLog.Close
Set arquivoLog = Nothing
Set fso = Nothing

End Sub

Public Sub leLog(t As Control)

Dim fso As New FileSystemObject

'declara as variáveis objetos
Dim arquivo As File
Dim fsoStream As TextStream
Dim strLinha As String
Dim arquivologerros As String
caminho = App.Path & "\sgrm.log"
'abre o arquivo para leitura
If fso.FileExists(caminho) Then
   Set arquivo = fso.GetFile(caminho)
   Set fsoStream = arquivo.OpenAsTextStream(ForReading)
Else
   MsgBox "O arquivo não existe", vbCritical
Exit Sub
End If

' le o arquivo linha a linha e exibe no text1
Do While Not fsoStream.AtEndOfStream
   strLinha = strLinha & fsoStream.ReadLine & vbCrLf
   t = strLinha
Loop

'libera as variaveis objeto
fsoStream.Close
Set fsoStream = Nothing
Set arquivo = Nothing
Set fso = Nothing

End Sub
