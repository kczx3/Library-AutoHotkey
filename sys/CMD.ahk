/*
    Ejecuta el o los comandos especificados en CMD.EXE.
    Parámetros:
        Commands:
            Los comandos a ejecutar separados por una nueva línea "`n".
        WorkingDir:
            La ruta del directorio de trabajo a utilizar. Por defecto utiliza el directorio de trabajo actual.
        Options:
            Wait = Espera a que los comandos terminen de ejecutarse y devuelve la salida.
            Hide = Oculta la ventana de CMD mientras se procesan los comandos.
             Min = Minimiza la ventana.
             Max = Maximiza la ventana.
            High = Cambia la prioridad del proceso a Alta.
        ProcessID:
            Devuelve el identificador del proceso de la ventana CMD.
    Return:
        Si se especificó la opción "W" devuelve el texto de salida, en caso contrario devuelve un objeto shell.exec.
    Ejemplos:
        MsgBox CMD("cd", A_ProgramFiles, "Wait", PID) . " [" . PID . "]"
        MsgBox CMD("tasklist",, "Wait Hide")
*/
CMD(Commands, WorkingDir := "", Options := "", ByRef ProcessID := "")
{
    Local shell := ComObjCreate("WScript.Shell")
        ,  exec := shell.Exec(A_ComSpec . " /Q /K echo off")    ; chcp 65001>null
    ProcessID := exec.ProcessId

    If (Options != "")
    {
        A_DetectHiddenWindows := TRUE
        WinWait("ahk_pid" . ProcessID)
        If (InStr(Options, "Hide"))
            WinHide("ahk_pid" . ProcessID)
        If (InStr(Options, "Min"))
            WinMinimize("ahk_pid" . ProcessID)
        If (InStr(Options, "Max"))
            WinMaximize("ahk_pid" . ProcessID)
        If (InStr(Options, "High"))
            ProcessSetPriority("High", ProcessID)
    }

    WorkingDir := DirExist(WorkingDir) ? WorkingDir : A_WorkingDir
    Commands := "cd /d " . RTrim(WorkingDir, "\") . "\`n" . Trim(Commands, "`s`t`r`n")
    exec.StdIn.WriteLine(Commands . "`nexit")

    Return InStr(Options, "Wait") ? Trim(exec.StdOut.ReadAll(), "`r`n") : exec
}
