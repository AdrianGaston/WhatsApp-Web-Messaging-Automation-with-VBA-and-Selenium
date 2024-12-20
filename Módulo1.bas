Attribute VB_Name = "Módulo1"
Option Explicit

Private cd As Selenium.ChromeDriver

Sub SendWhats()
    'Declaração de variáveis
    Dim cxPesquisa As WebElement
    Dim cxMensagem As WebElement
    Dim localMsg As New Keys
    
    'Inicializa o ChromeDriver
    Set cd = New Selenium.ChromeDriver
    
'Tratamento de erro
On Error GoTo TratarErro

    With cd
        .SetBinary "C:\Program Files\Google\Chrome\Application\chrome.exe" 'Define o caminho para o executável do Chrome
        .SetProfile Environ("LOCALAPPDATA") & "\Google\Chrome\User Data\Default" 'Define o caminho do perfil do Chrome onde a conta está logada
        '.AddArgument "--remote-debugging-port=9222" 'Argumento para depuração remota (permite reutilizar sessões logadas)
        .AddArgument "--start-maximized" 'Inicia a janela maximizada
        .AddArgument "--hide-crash-restore-bubble" 'Evita que o Chrome exiba a mensagem de 'restauração de sessão'
        .AddArgument "--disable-notifications" 'Desabilita as notificações do navegador
        .Start 'Inicia o Chrome
        .Get "https://web.whatsapp.com/" 'Acessa o WhatsApp Web
        .Wait 10000 'Espera de 10 segundos
    End With

    cd.Timeouts.PageLoad = 60000 'Tempo máximo para carregar a página (60 segundos)
    cd.Timeouts.ImplicitWait = 60000 'Tempo máximo para localizar o elemento (60 segundos)
    
    'Declara as variáveis do Timer
    Dim tempoInicial As Single, tempoLimite As Single
    
    tempoInicial = Timer
    tempoLimite = tempoInicial + 60 'Tempo máximo de 1 minuto

    'Loop para aguardar a página ser carregada completamente
    Do While cxPesquisa Is Nothing And Timer < tempoLimite
        On Error Resume Next
        Set cxPesquisa = cd.FindElementByXPath("//*[@id=""side""]/div[1]/div/div[2]/div[2]/div/div/p")
        On Error GoTo 0
        cd.Wait 1000
    Loop

    'Restaura o tratamento normal
    On Error GoTo TratarErro

    If cxPesquisa Is Nothing Then
        MsgBox "Não foi possível carregar o WhatsApp. Tente novamente mais tarde.", vbCritical, "Erro de Carregamento"
        Exit Sub
    End If
    
    'Encontra o campo de pesquisa
    Set cxPesquisa = cd.FindElementByXPath("//*[@id=""side""]/div[1]/div/div[2]/div[2]/div/div/p")
    cxPesquisa.SendKeys "Saved Messages"
    cd.Wait 1000 'Espera

    'Pressiona Enter para selecionar o contato
    cxPesquisa.SendKeys localMsg.Enter
    cd.Wait 1000 'Espera

    'Localiza o campo de mensagem
    Set cxMensagem = cd.FindElementByXPath("//*[@id=""main""]/footer/div[1]/div/span/div/div[2]/div[1]/div/div[1]/p")
    
    'Envia a mensagem
    cxMensagem.SendKeys "Hello World!"
    cd.Wait 1000 'Espera

    'Pressiona Enter para enviar a mensagem
    cxMensagem.SendKeys localMsg.Enter
    cd.Wait 3000 'Espera
    
    'Mensagem de conexão bem-sucedida no console
    Debug.Print "Conexão estabelecida com sucesso!"
    
    Exit Sub 'Sai antes do tratamento de erros
    
TratarErro:
        Debug.Print "Erro inesperado: " & Err.Description & " (" & Err.Number & ")"
        MsgBox "Erro inesperado: " & Err.Description & " (" & Err.Number & ")", vbCritical, "Erro de Conexão"
    
End Sub
