/*
    Muestra un diálogo para seleccionar una carpeta.
    Parámetros:
        Owner / Title:
            El identificador de la ventana propietaria de este diálogo. Este valor puede ser cero.
            Un Array con el identificador de la ventana propietaria y el título. Si el título es una cadena vacía se establece al por defecto "Seleccionar carpeta".
        StartingFolder:
            La ruta al directorio seleccioado por defecto. Si el directorio no existe, busca en directorios superiores.
        CustomPlaces:
            Especifica un Array con los directorios personalizados que se mostrarán en el panel izquierdo. Los directorios inexistentes serán omitidos.
            Para especificar la hubicación en la lista, especifique un Array con el directorio y la hubicación de este (0 = Inferior, 1 = Superior).
            Generalmente se suelte utilizar con la opción FOS_HIDEPINNEDPLACES.
        Options:
            Determina el comportamiento del diálogo. Este parámetro debe ser uno o más de los siguientes valores.
            0x00000200 (FOS_ALLOWMULTISELECT) = Permite seleccionar más de un directorio.
            0x00040000 (FOS_HIDEPINNEDPLACES) = Ocultar elementos que se muestran de forma predeterminada en el panel de navegación de la vista.
            0x02000000  (FOS_DONTADDTORECENT) = No agregua el elemento que se abre o guarda en la lista de documentos recientes (función SHAddToRecentDocs).
            0x10000000  (FOS_FORCESHOWHIDDEN) = Incluye elementos ocultos y del sistema.
            Puede consultar todos los valores disponibles en https://msdn.microsoft.com/en-us/library/windows/desktop/dn457282(v=vs.85).aspx.
    Return:
        Devuelve 0 si el usuario canceló el diálogo, caso contrario devuelve el directorio seleccionado. El directorio nunca termina con "\".
        Si especificó la opción FOS_ALLOWMULTISELECT y la función tuvo éxito, devuelve un Array con los directorios seleccionados.
    Ejemplo:
        Result := ChooseFolder( [0, "Título del diálogo - Seleccionar carpeta.."]
                              , A_WinDir . "\System32\Non-existent folder\Carpeta inexistente"
                              , [A_WinDir,A_Desktop,A_Temp,A_Startup,A_ProgramFiles]
                              , 0x10000000 | 0x02000000 | 0x00000200 )
        If ((List := "") == "" && Result != FALSE)
            For Each, Directory in Result
                List .= Directory . "`n"
        MsgBox List
*/
ChooseFolder(Owner, StartingFolder := "", CustomPlaces := "", Options := 0)
{
    ; IFileOpenDialog interface
    ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb775834(v=vs.85).aspx
    local IFileOpenDialog := ComObjCreate("{DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7}", "{D57C7288-D4AD-4768-BE02-9D969532D960}")
        ,           Title := IsObject(Owner) ? Owner[2] . "" : ""
        ,           Flags := 0x20 | Options    ; FILEOPENDIALOGOPTIONS enumeration (https://msdn.microsoft.com/en-us/library/windows/desktop/dn457282(v=vs.85).aspx)
        ,      IShellItem := PIDL := 0         ; PIDL recibe la dirección de memoria a la estructura ITEMIDLIST que debe ser liberada con la función CoTaskMemFree
        ,             Obj := {}, foo := bar := ""
    Owner := IsObject(Owner) ? Owner[1] : (WinExist("ahk_id" . Owner) ? Owner : 0)
    CustomPlaces := IsObject(CustomPlaces) || CustomPlaces == "" ? CustomPlaces : [CustomPlaces]


    while (InStr(StartingFolder, "\") && !DirExist(StartingFolder))    ; si el directorio no existe buscamos directorios superiores
        StartingFolder := SubStr(StartingFolder, 1, InStr(StartingFolder, "\",, -1) - 1)
    if ( DirExist(StartingFolder) )
    {
        DllCall("Shell32.dll\SHParseDisplayName", "UPtr", &StartingFolder, "Ptr", 0, "UPtrP", PIDL, "UInt", 0, "UInt", 0)
        DllCall("Shell32.dll\SHCreateShellItem", "Ptr", 0, "Ptr", 0, "UPtr", PIDL, "UPtrP", IShellItem)
        ObjRawSet(Obj, IShellItem, PIDL)    ; guardamos la interfaz «IShellItem» y la estructura «PIDL» para liberarlas al final
        ; IFileDialog::SetFolder method
        ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb761828(v=vs.85).aspx
        DllCall(NumGet(NumGet(IFileOpenDialog)+12*A_PtrSize), "Ptr", IFileOpenDialog, "UPtr", IShellItem)
    }


    if ( IsObject(CustomPlaces) )
    {
        local Directory := ""
        For foo, Directory in CustomPlaces    ; foo = index
        {
            foo := IsObject(Directory) ? Directory[2] : 0    ; FDAP enumeration (https://msdn.microsoft.com/en-us/library/windows/desktop/bb762502(v=vs.85).aspx)
            if ( DirExist(Directory := IsObject(Directory) ? Directory[1] : Directory) )
            {
                DllCall("Shell32.dll\SHParseDisplayName", "UPtr", &Directory, "Ptr", 0, "UPtrP", PIDL, "UInt", 0, "UInt", 0)
                DllCall("Shell32.dll\SHCreateShellItem", "Ptr", 0, "Ptr", 0, "UPtr", PIDL, "UPtrP", IShellItem)
                ObjRawSet(Obj, IShellItem, PIDL)
                ; IFileDialog::AddPlace method
                ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb775946(v=vs.85).aspx
                DllCall(NumGet(NumGet(IFileOpenDialog)+21*A_PtrSize), "UPtr", IFileOpenDialog, "UPtr", IShellItem, "UInt", foo)
            }
        }
    }


    ; IFileDialog::SetTitle method
    ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb761834(v=vs.85).aspx
    DllCall(NumGet(NumGet(IFileOpenDialog)+17*A_PtrSize), "UPtr", IFileOpenDialog, "UPtr", Title == "" ? 0 : &Title)

    ; IFileDialog::SetOptions method
    ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb761832(v=vs.85).aspx
    DllCall(NumGet(NumGet(IFileOpenDialog)+9*A_PtrSize), "UPtr", IFileOpenDialog, "UInt", Flags)


    ; IModalWindow::Show method
    ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb761688(v=vs.85).aspx
    local Result := []
    if ( !DllCall(NumGet(NumGet(IFileOpenDialog)+3*A_PtrSize), "UPtr", IFileOpenDialog, "Ptr", Owner, "UInt") )
    {
        ; IFileOpenDialog::GetResults method
        ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb775831(v=vs.85).aspx
        local IShellItemArray := 0    ; IShellItemArray interface (https://msdn.microsoft.com/en-us/library/windows/desktop/bb761106(v=vs.85).aspx)
        DllCall(NumGet(NumGet(IFileOpenDialog)+27*A_PtrSize), "UPtr", IFileOpenDialog, "UPtrP", IShellItemArray)

        ; IShellItemArray::GetCount method
        ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb761098(v=vs.85).aspx
        local Count := 0    ; pdwNumItems
        DllCall(NumGet(NumGet(IShellItemArray)+7*A_PtrSize), "UPtr", IShellItemArray, "UIntP", Count)

        local Buffer := ""
        loop ( 0*VarSetCapacity(Buffer,32767 * 2) + Count )
        {
            ; IShellItemArray::GetItemAt method
            ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb761100(v=vs.85).aspx
            DllCall(NumGet(NumGet(IShellItemArray)+8*A_PtrSize), "UPtr", IShellItemArray, "UInt", A_Index-1, "UPtrP", IShellItem)
            DllCall("Shell32.dll\SHGetIDListFromObject", "UPtr", IShellItem, "UPtrP", PIDL)
            DllCall("Shell32.dll\SHGetPathFromIDListEx", "UPtr", PIDL, "Str", Buffer, "UInt", 32767, "UInt", 0)
            ObjRawSet(Obj, IShellItem, PIDL), ObjPush(Result, RTrim(Buffer, "\"))
        }

        ObjRelease(IShellItemArray)
    }


    for foo, bar in Obj    ; foo = IShellItem interface (ptr)  |  bar = PIDL structure (ptr)
        ObjRelease(foo), DllCall("Ole32.dll\CoTaskMemFree", "UPtr", bar)
    ObjRelease(IFileOpenDialog)

    return ObjLength(Result) ? ( Options & 0x200 ? Result : Result[1] ) : FALSE
}
