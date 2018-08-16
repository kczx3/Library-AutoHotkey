/*
    Ejecuta un Script en un nuevo proceso AutoHotkey y devuelve un objeto.
    Parámetros:
        Script : El Script a ejecutar.
        Params : Los parámetros. Puede ser un Array o una cadena.
        AhkPath: La ruta a AutoHotkey.exe. Si no se especifica, busca en el directorio actual del Script, si no lo encuentra busca en A_AhkPath.
    Return:
        Devuelve un objeto de clase __Class_ExecScript si tuvo éxito, caso contrario devuelve 0.
        Cuando el objeto es eliminado, el proceso termina. Para evitar esto, llamar al método AutoTerminate pasando FALSE.
    Ejemplo 1:
        Script1 := ExecScript("MsgBox('Ejecutando desde Script #1')")
        Script2 := ExecScript("MsgBox('Ejecutando desde Script #2')")

        MsgBox("Ejecutando desde Script Principal")
        Script1 := Script2 := ""
        MsgBox("ok!")
        ExitApp
    Ejemplo 2:
        ; Nuevo proceso
        Script := "MsgBox('Nuevo Proceso...')`n"                                    ;Mostramos un mensaje.
                . "MsgBox('Memoria compartida (nuevo proceso):``n' . SM_Read())`n"  ;Leemos la memoria compartida.
                . "SM_Write('Finalizando...')`n"                                    ;Escribimos en la memoria compartida (debemos asegurarnos de que hay espacio).
                . "ExitApp"                                                         ;Terminamos el nuevo proceso.
        
        ; Este proceso
        Script := ExecScript(Script) ;Iniciamos el nuevo proceso.
        Script.Realloc(14)           ;Ajustamos la capacidad de la memoria compartida para almacenar hasta 14 caracteres en UTF-16.
        Script.Write("Hola Mundo!")  ;Escribimos una cadena en la memoria compartida, que podrá ser recuperada desde el proceso mediante la función SM_Read().

        Script.Wait()                                                                                   ;Esperamos a que termine el proceso iniciado.
        MsgBox("Memoria compartida (proceso principal):`n" . Script.Read() . "`nBytes: " . Script.Size) ;Leemos la memoria compartida.

        Script := "" ;Liberamos la memoria.
        ExitApp      ;Terminamos el Script.
    Ejemplo 3:
        Script := ExecScript()
        Script.Write("Hola Mundo!")

        MsgBox(Script.Read())

        Script.RawRead(Data)
        MsgBox(StrGet(&Data, "UTF-16"))
        Data := ""

        ExitApp
    Ejemplo 4:
        Script := ExecScript()

        VarSetCapacity(Data, Size := 4 * 2 + 2)
        StrPut("Hola", &Data, 4, "UTF-16")
        NumPut(0x0000, Data, 8, "UShort")

        Script.RawWrite(&Data, Size)

        MsgBox(Script.Read())
        ExitApp
    Ejemplo 5:
        Script := ExecScript("MsgBox()`nSM_RawRead(Data)`nMsgBox(StrGet(&Data, 'UTF-16'))`nExitApp")
        Script.Write("Hola Mundo!")
        Script.Wait()
        ExitApp
*/
ExecScript(Script := "", Params := "", AhkPath := "")
{
    If (!FileExist(AhkPath) && !FileExist(AhkPath := A_WorkingDir . "\AutoHotkey.exe") && !FileExist(AhkPath := A_AhkPath))
        Return (FALSE)

    If (IsObject(Params))
        For Each, Param in Params
            Params .= Chr(34) . StrReplace(Param, Chr(34)) Chr(34) . A_Space

    FMName  := "AHK-" . A_TickCount

    Script0 := "`n#SingleInstance Off   `n#Persistent"
    Script0 .= "`nGlobal __SM_Size"
    Script0 .= "`nOnMessage(0x5555, 'SCRIPT_EXIT')"
    Script0 .= "`nOnMessage(0x5556, 'SCRIPT_SETSIZE')"
    Script0 .= "`n" . Script . "`nReturn"
    Script0 .= "`nSCRIPT_EXIT(wParam, lParam) {`nExitApp`n}"
    Script0 .= "`nSCRIPT_SETSIZE(wParam, lParam) {`n__SM_Size := wParam`n}"
    Script0 .= "`nSM_Read() {"
    Script0 .= "`nhMapFile := DllCall('Kernel32.dll\OpenFileMappingW', 'UInt', 0xF001F, 'Int', FALSE, 'Str', '" . FMName . "', 'Ptr')"
    Script0 .= "`npBuffer := DllCall('Kernel32.dll\MapViewOfFile', 'Ptr', hMapFile, 'UInt', 0xF001F, 'UInt', 0, 'UInt', 0, 'UPtr', 0, 'Ptr')"
    Script0 .= "`nString := StrGet(pBuffer, 'UTF-16')"
    Script0 .= "`nDllCall('Kernel32.dll\UnmapViewOfFile', 'Ptr', pBuffer)"
    Script0 .= "`nDllCall('Kernel32.dll\CloseHandle', 'Ptr', hMapFile)"
    Script0 .= "`nReturn (String)   `n}"
    Script0 .= "`nSM_Write(String) {"
    Script0 .= "`nhMapFile := DllCall('Kernel32.dll\OpenFileMappingW', 'UInt', 0xF001F, 'Int', FALSE, 'Str', '" . FMName . "', 'Ptr')"
    Script0 .= "`npBuffer := DllCall('Kernel32.dll\MapViewOfFile', 'Ptr', hMapFile, 'UInt', 0xF001F, 'UInt', 0, 'UInt', 0, 'UPtr', 0, 'Ptr')"
    Script0 .= "`nStrPut(String, pBuffer, Size := StrLen(String), 'UTF-16')"
    Script0 .= "`nNumPut(0x0000, pBuffer, Size * 2, 'UShort')"
    Script0 .= "`nDllCall('Kernel32.dll\UnmapViewOfFile', 'Ptr', pBuffer)"
    Script0 .= "`nDllCall('Kernel32.dll\CloseHandle', 'Ptr', hMapFile)"
    Script0 .= "`n}"
    Script0 .= "`nSM_RawWrite(Address, Size) {"
    Script0 .= "`nhMapFile := DllCall('Kernel32.dll\OpenFileMappingW', 'UInt', 0xF001F, 'Int', FALSE, 'Str', '" . FMName . "', 'Ptr')"
    Script0 .= "`npBuffer := DllCall('Kernel32.dll\MapViewOfFile', 'Ptr', hMapFile, 'UInt', 0xF001F, 'UInt', 0, 'UInt', 0, 'UPtr', 0, 'Ptr')"
    Script0 .= "`nDllCall('msvcrt.dll\memcpy_s', 'UPtr', pBuffer, 'UPtr', Size, 'UPtr', Address, 'UPtr', Size, 'Cdecl')"
    Script0 .= "`nDllCall('Kernel32.dll\UnmapViewOfFile', 'Ptr', pBuffer)"
    Script0 .= "`nDllCall('Kernel32.dll\CloseHandle', 'Ptr', hMapFile)"
    Script0 .= "`n}"
    Script0 .= "`nSM_RawRead(ByRef Buffer) {"
    Script0 .= "`nhMapFile := DllCall('Kernel32.dll\OpenFileMappingW', 'UInt', 0xF001F, 'Int', FALSE, 'Str', '" . FMName . "', 'Ptr')"
    Script0 .= "`npBuffer := DllCall('Kernel32.dll\MapViewOfFile', 'Ptr', hMapFile, 'UInt', 0xF001F, 'UInt', 0, 'UInt', 0, 'UPtr', 0, 'Ptr')"
    Script0 .= "`nVarSetCapacity(Buffer, __SM_Size)"
    Script0 .= "`nDllCall('msvcrt.dll\memcpy_s', 'UPtr', &Buffer, 'UPtr', __SM_Size, 'UPtr', pBuffer, 'UPtr', __SM_Size, 'Cdecl')"
    Script0 .= "`nDllCall('Kernel32.dll\UnmapViewOfFile', 'Ptr', pBuffer)"
    Script0 .= "`nDllCall('Kernel32.dll\CloseHandle', 'Ptr', hMapFile)"
    Script0 .= "`n}"
    
    Loop
        If (File := FileOpen(F := A_Temp . "\~tmp-ahk" . A_Index, "w-rwd", "UTF-16"))
            Break

    File.Write(Script0)
    File.Close()

    Run(Chr(34) . AhkPath . Chr(34) . A_Space . Chr(34) . F . Chr(34) . A_Space . Params,,, PID)

    Return (PID ? New __Class_ExecScript(PID, FMName) : 0)
}




Class __Class_ExecScript
{
    hWnd        := 0
    ProcessId   := 0
    hProcess    := 0
    AT          := TRUE
    hMapFile    := 0
    Size        := 0
    FMName      := ""

    __New(ProcessId, FMName, Size := 2)
    {
        DetectHiddenWindows('On')
        WinWait("ahk_pid" . ProcessId,, 3)
       
        This.hWnd      := WinExist()
        This.ProcessId := ProcessId
        This.hProcess  := DllCall("Kernel32.dll\OpenProcess", "UInt", 0x1F0FFF, "Int", FALSE, "UInt", ProcessId, "Ptr")

        If (!This.hProcess || !This.hWnd)
            Return (FALSE)

        ; https://msdn.microsoft.com/en-us/library/aa366537(v=vs.85).aspx
        This.FMName   := FMName
        This.Size     := Size
        This.hMapFile := DllCall("Kernel32.dll\CreateFileMappingW", "Ptr", 0, "Ptr", 0, "UInt", 0x40, "UInt", 0, "UInt", This.Size, "Str", This.FMName, "Ptr")

        ; https://msdn.microsoft.com/en-us/library/aa366761(v=vs.85).aspx
        pBuffer       := DllCall("Kernel32.dll\MapViewOfFile", "Ptr", This.hMapFile, "UInt", 0xF001F, "UInt", 0, "UInt", 0, "UPtr", 2, "Ptr")
        NumPut(0x0000, pBuffer, 0, "UShort")
        DllCall("Kernel32.dll\UnmapViewOfFile", "Ptr", pBuffer)

        DllCall("User32.dll\PostMessageW", "Ptr", This.hWnd, "UInt", 0x5556, "Ptr", This.Size, "Ptr", 0)
    }

    __Delete()
    {
        If (This.AT)
        {
            DllCall("User32.dll\PostMessageW", "Ptr", This.hWnd, "UInt", 0x5555, "Ptr", 0, "Ptr", 0)
            WinWaitClose("ahk_id" . This.hWnd,, 3)
            If (ErrorLevel)
                DllCall("Kernel32.dll\TerminateProcess", "Ptr", This.hProcess)
        }

        DllCall("Kernel32.dll\CloseHandle", "Ptr", This.hProcess)
        DllCall("Kernel32.dll\CloseHandle", "Ptr", This.hMapFile)
    }

    AutoTerminate(State)
    {
        This.AT := !!State
    }

    Wait()
    {
        ProcessWaitClose(This.ProcessId)
    }

    Realloc(Size, Mode := 0) ;Mode: 0=Char|1=Bytes
    {
        DllCall("Kernel32.dll\CloseHandle", "Ptr", This.hMapFile)

        This.Size     := Mode ? Size : Size * 2 + 2
        This.hMapFile := DllCall("Kernel32.dll\CreateFileMappingW", "Ptr", 0, "Ptr", 0, "UInt", 0x40, "UInt", 0, "UInt", This.Size, "Str", This.FMName, "Ptr")
        
        DllCall("User32.dll\PostMessageW", "Ptr", This.hWnd, "UInt", 0x5556, "Ptr", This.Size, "Ptr", 0)
    }

    Write(String, Size := 0)
    {
        String        .= ""
        Size          := Size ? Size : StrLen(String)

        If (Size * 2 + 2 > This.Size)
        {
            DllCall("Kernel32.dll\CloseHandle", "Ptr", This.hMapFile)

            This.Size     := Size * 2 + 2
            This.hMapFile := DllCall("Kernel32.dll\CreateFileMappingW", "Ptr", 0, "Ptr", 0, "UInt", 0x40, "UInt", 0, "UInt", This.Size, "Str", This.FMName, "Ptr")
        
            DllCall("User32.dll\PostMessageW", "Ptr", This.hWnd, "UInt", 0x5556, "Ptr", This.Size, "Ptr", 0)
        }

        pBuffer       := DllCall("Kernel32.dll\MapViewOfFile", "Ptr", This.hMapFile, "UInt", 0xF001F, "UInt", 0, "UInt", 0, "UPtr", Size * 2 + 2, "Ptr")
        StrPut(String, pBuffer, Size, "UTF-16")
        NumPut(0x0000, pBuffer, Size * 2, "UShort")
        DllCall("Kernel32.dll\UnmapViewOfFile", "Ptr", pBuffer)
    } ;https://msdn.microsoft.com/en-us/library/aa366761(v=vs.85).aspx

    RawWrite(Address, Size)
    {
        If (Size > This.Size)
        {
            DllCall("Kernel32.dll\CloseHandle", "Ptr", This.hMapFile)

            This.Size     := Size
            This.hMapFile := DllCall("Kernel32.dll\CreateFileMappingW", "Ptr", 0, "Ptr", 0, "UInt", 0x40, "UInt", 0, "UInt", This.Size, "Str", This.FMName, "Ptr")
            
            DllCall("User32.dll\PostMessageW", "Ptr", This.hWnd, "UInt", 0x5556, "Ptr", This.Size, "Ptr", 0)
        }

        pBuffer       := DllCall("Kernel32.dll\MapViewOfFile", "Ptr", This.hMapFile, "UInt", 0xF001F, "UInt", 0, "UInt", 0, "UPtr", Size * 2 + 2, "Ptr")
        
        DllCall("msvcrt.dll\memcpy_s", "UPtr", pBuffer, "UPtr", Size, "UPtr", Address, "UPtr", Size, "Cdecl")

        DllCall("Kernel32.dll\UnmapViewOfFile", "Ptr", pBuffer)
    }

    Read()
    {
        pBuffer := DllCall("Kernel32.dll\MapViewOfFile", "Ptr", This.hMapFile, "UInt", 0xF001F, "UInt", 0, "UInt", 0, "UPtr", 0, "Ptr")
        String  := StrGet(pBuffer, "UTF-16")
        DllCall("Kernel32.dll\UnmapViewOfFile", "Ptr", pBuffer)

        Return (String)
    }

    RawRead(ByRef Buffer)
    {
        pBuffer := DllCall("Kernel32.dll\MapViewOfFile", "Ptr", This.hMapFile, "UInt", 0xF001F, "UInt", 0, "UInt", 0, "UPtr", 0, "Ptr")
        
        VarSetCapacity(Buffer, This.Size)
        DllCall("msvcrt.dll\memcpy_s", "UPtr", &Buffer, "UPtr", This.Size, "UPtr", pBuffer, "UPtr", This.Size, "Cdecl")

        DllCall("Kernel32.dll\UnmapViewOfFile", "Ptr", pBuffer)
    }
}
