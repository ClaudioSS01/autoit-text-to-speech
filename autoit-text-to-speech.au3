#include <MsgBoxConstants.au3>

Func SpeakTextUsingVBScript($text)
    Local $vbscriptCode = 'CreateObject("SAPI.SpVoice").Speak("' & $text & '") : window.close'
    Local $vbscriptFile = @TempDir & "\speak.vbs"
    
    ; Escrever o código VBScript em um arquivo temporário
    FileWrite($vbscriptFile, $vbscriptCode)
    
    ; Executar o código VBScript usando mshta
    Run(@ComSpec & ' /c mshta vbscript:Execute("CreateObject(""Scripting.FileSystemObject"").GetFile(WScript.Arguments(0)).OpenAsTextStream(1, -2).ReadAll():Close")("' & $vbscriptFile & '")(close)')
    
    ; Aguardar até que a fala seja concluída (pode ajustar conforme necessário)
    Sleep(3000)
    
    ; Excluir o arquivo temporário
    FileDelete($vbscriptFile)
EndFunc

; Exemplo de uso da função
Local $inputText = InputBox("Texto para Falar", "Digite o texto que deseja que seja falado:")
If @error = 0 And $inputText <> "" Then
    SpeakTextUsingVBScript($inputText)
EndIf
