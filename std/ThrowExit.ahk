/*
    Muestra un mensaje de error y termina el Script.
    Parámetros:
        Message : El mensaje de error.
        ExitCode: El código de salida para este proceso.
*/
ThrowExit(Message, ExitCode := 0)
{
    static cv := [], tp := cv
    if (tp != cv)
        Throw tp
    tp := Message
    DllCall(CallbackCreate(A_ThisFunc))
    Exitapp ExitCode
} ; https://autohotkey.com/boards/viewtopic.php?f=6&t=39604&p=181047
