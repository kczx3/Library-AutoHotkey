/*
    Ejecuta el código especificado sin tocar el sistema de archivos (no crea ni modifica ningún archivo).
    Parámetros:
        Script    : El código AHK a ejecutar. La codificación utilizada es UTF-16.
        AhkPath   : La ruta a AHK y los parámetros a pasar. Este parámetro es opcional. Si AHK se encuentra en el directorio actual, puede especificar solo "AutoHotkey.exe Params".
        WorkingDir: El directorio de trabajo actual del script. Si expecifica una cadena vacía, hereda el directorio actual.
    Return:
        Si tuvo éxito devuelve el identificador del proceso.
    Ejemplo #1:
        ExecScript("MsgBox(A_Args[1] . A_Args[2] . A_Args[3])", Chr(34) . A_AhkPath . "`" Hola `" Mundo`" !")
    Ejemplo #2:
        Script_1 := ExecScript("MsgBox `"Script_1`"")
        Script_2 := ExecScript("MsgBox `"Script_2`"")
        ProcessWaitClose(Script_1), ProcessWaitClose(Script_2)
        MsgBox "Echo!"
*/
ExecScript(Script, AhkPath := "", WorkingDir := "")
{
    If ((Script := Trim(Script)) == "")
        Return 0
    Script := "A_WorkingDir := `"" . (DirExist(WorkingDir) ? WorkingDir : A_WorkingDir) . "`"`n" . Script

    Local Args
    SplitArgs(AhkPath, Args), AhkPath := AhkPath == "" || InStr(AhkPath, ":") ? AhkPath : A_WorkingDir . "\" . AhkPath
    AhkPath := DirExist(AhkPath) || !FileExist(AhkPath) ? A_AhkPath : AhkPath
    If (DirExist(AhkPath) || !FileExist(AhkPath))
        Return 0

    Local     Pipe := [0, 0]
        , PipeName := "\\.\pipe\AHK_" . A_TickCount
    Loop 2
        ; https://msdn.microsoft.com/en-us/library/windows/desktop/aa365150(v=vs.85).aspx
        Pipe[A_Index] := DllCall("Kernel32.dll\CreateNamedPipeW", "UPtr", &PipeName    ; lpName
                                                                , "UInt", 2            ; dwOpenMode           -->     PIPE_ACCESS_OUTBOUND
                                                                , "UInt", 0            ; dwPipeMode           -->           PIPE_TYPE_BYTE
                                                                , "UInt", 255          ; nMaxInstances        --> PIPE_UNLIMITED_INSTANCES
                                                                , "UInt", 0            ; nOutBufferSize
                                                                , "UInt", 0            ; nInBufferSize
                                                                , "UInt", 0            ; nDefaultTimeOut      --> default time-out of 50 milliseconds
                                                                , "UPtr", 0, "Ptr")    ; lpSecurityAttributes --> NULL (default security descriptor and the handle cannot be inherited)          

    Local ProcessId
    Run(Chr(34) . AhkPath . "`" /CP1200 " . PipeName . " " . Args,,, ProcessId)

    ; https://msdn.microsoft.com/en-us/library/windows/desktop/aa365146(v=vs.85).aspx
    DllCall("Kernel32.dll\ConnectNamedPipe", "Ptr", Pipe[1], "UPtr", 0)
    ; https://msdn.microsoft.com/en-us/library/windows/desktop/ms724211(v=vs.85).aspx
    DllCall("Kernel32.dll\CloseHandle", "Ptr", Pipe[1])
    DllCall("Kernel32.dll\ConnectNamedPipe", "Ptr", Pipe[2], "UPtr", 0)

    FileOpen(Pipe[2], "h", "UTF-16").Write(Script)
    DllCall("Kernel32.dll\CloseHandle", "Ptr", Pipe[2])

    Return ProcessId


    SplitArgs(ByRef AhkPath, ByRef Args)
    {
        If (SubStr(AhkPath := Trim(AhkPath), 1, 1) == Chr(34))
        {
            Local n := InStr(AhkPath, Chr(34),, 2)
            Args    := n ? Trim(SubStr(AhkPath, n+2)) : ""
            AhkPath := Trim(SubStr(n ? SubStr(AhkPath, 1, n-1) : AhkPath, 2))
        }
        Else
        {
            Local n := InStr(AhkPath, A_Space)
            Args    := SubStr(AhkPath, n+1)
            AhkPath := SubStr(AhkPath, 1, n-1)
        }
    }
} ; CREDITS TO cocobelgica (https://autohotkey.com/boards/viewtopic.php?f=6&t=5090)





/*
    Termina el Script especificado.
    Parámetros:
        ProcessId: El identificador del proceso del Script a terminar.
*/
ExitScript(ProcessId)
{
    PostMessage(0x111, 65307, 0,, "ahk_pid" . ProcessId)
    If (ProcessWaitClose(ProcessId, 5))
        ProcessClose(ProcessId)
}
