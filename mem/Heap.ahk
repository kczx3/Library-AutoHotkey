/* EXAMPLE #1
Str := "Hello World!"

Mem := new Heap

Mem.Alloc( StrLen(Str) )

StrPut(Str, Mem.Ptr, Mem.Size, "UTF-8")

MsgBox StrGet(Mem.Ptr, Mem.Size, "UTF-8")
*/


/* EXAMPLE #2
Mem := new Heap().Alloc(90 - 64)

Loop (Mem.Size)
    NumPut(64 + A_Index, Mem.Ptr + (A_Index - 1), "UChar")

MsgBox StrGet(Mem.Ptr, Mem.Size, "UTF-8")

Ptr := Mem.Ptr
Loop (Mem.Size)
    Ptr := NumPut(91 - A_Index, Ptr, "UChar")

MsgBox StrGet(Mem.Ptr, Mem.Size, "UTF-8")
*/


/* EXAMPLE #3
Str := "A B C"

Mem := new Heap().Alloc(StrLen(Str) * 2)

Mem.CopyFrom(&Str, 0, Mem.Size)

MsgBox StrGet(Mem.Ptr, StrLen(Str), "UTF-16")    ; StrLen(Str) = Mem.Size//2
*/


/* EXAMPLE #4
MsgBox NumGet((new Heap().Alloc(1).Fill(255)).Ptr, "UChar")
*/

/* EXAMPLE #5
Str := "Hello World!"
Mem := new Heap().Alloc(StrLen(Str)*2+2)
MsgBox "Str: " . Str . " (" . StrLen(Str) .  ")`nStr Length: " . Mem.CopyFrom(&Str, 0, Mem.Size).Length
*/


/* EXAMPLE #6
Mem := new Heap().Alloc(24)
StrPut("Hello World!", Mem.Ptr, 12, "UTF-16")
;      ----------------- World ------------------   -------------   ---------------   ------------------ Hello ----------------   ------------------ ! --------------------
MsgBox StrGet(Mem.Clone(12, 10).Ptr, 5, "UTF-16") . StrGet(Mem.Ptr+10, 1, "UTF-16") . StrGet(Mem.Clone(0, 10).Ptr, 5, "UTF-16") . StrGet(Mem.Clone(22, 2).Ptr, 1, "UTF-16")
*/


/* EXAMPLE #7
Mem := new Heap().Alloc(26)
StrPut("Hello World!", Mem.Ptr, "UTF-16")
Mem.Trim(12, 10)    ; World | Size=10
MsgBox StrGet(Mem.Ptr, 5, "UTF-16")
*/


/* EXAMPLE #8
Mem := new Heap().Alloc(1024)
StrPut("XXX |Hello World!| XXX", Mem.Ptr, "UTF-8")    ; puede probar reemplazando 'XXX' por cualquier otra cadena
Mem.Trim(Mem.Chr(Ord("|"))-Mem.Ptr+1).Trim(0, Mem.Chr(Ord("|"), 1)-Mem.Ptr)
MsgBox StrGet(Mem.Ptr, Mem.Size, "UTF-8")
*/




/*
    Clase para reservar y utilizar memoria del montón (heap).
    La dirección de memoria se almacena en Heap->Ptr.
    Para poder reservar memoria debe llamar primero a Heap::Alloc, luego debe utilizar Heap::ReAlloc en caso de querer modificarse.
    Si utiliza Heap::Free, para volver a utilizar la memoria debe reservarla nuevamente utilizando Heap::Alloc.
    No debe utilizar Heap::Alloc si ya reservó memoria, y no debe utilizar Heap::ReAlloc si aún no se reservó memoria.
    No se comprueba ninguno de los parámetros pasados por motivos de rendimiento, por lo que debe asegurarse de pasar los valores correctos o la aplicación podría dejar de funcionar.
*/
Class Heap    ; https://msdn.microsoft.com/en-us/library/windows/desktop/aa366711(v=vs.85).aspx
{
    /*
        Crea un identificador de memoria Heap.
        HEAP_NO_SERIALIZE            0x00000001
        HEAP_GENERATE_EXCEPTIONS     0x00000004
        HEAP_CREATE_ENABLE_EXECUTE   0x00040000
    */
    __New(Options := 4, InitialSize := 0, MaximumSize := 0)
    {
        ; HeapCreate function
        ; https://msdn.microsoft.com/en-us/library/windows/desktop/aa366599(v=vs.85).aspx
        ObjRawSet(this, "hHeap", DllCall("Kernel32.dll\HeapCreate", "UInt", Options, "UPtr", InitialSize, "UPtr", MaximumSize, "Ptr"))
        ObjRawSet(this, "Ptr", 0)
    }

    /*
        Al eliminar el objeto Heap se libera la memoria reservada y se elimina el identificador Heap.
    */
    __Delete()
    {
        ; HeapDestroy function
        ; https://msdn.microsoft.com/en-us/library/windows/desktop/aa366700(v=vs.85).aspx
        DllCall("Kernel32.dll\HeapDestroy", "Ptr", this.hHeap)
        ; Processes can call HeapDestroy without first calling the HeapFree function to free memory allocated from the heap.
    }

    /*
        Reserva la cantidad de bytes especificados de memoria. Una vez reservada, debe utilizar Heap::ReAlloc para modificar la cantidad de memoria reservada.
        HEAP_NO_SERIALIZE          0x00000001
        HEAP_GENERATE_EXCEPTIONS   0x00000004
        HEAP_ZERO_MEMORY           0x00000008
    */
    Alloc(Bytes, Flags := 4)
    {
        ; HeapReAlloc function
        ; https://msdn.microsoft.com/en-us/library/windows/desktop/aa366597(v=vs.85).aspx
        ObjRawSet(this, "Ptr", DllCall("Kernel32.dll\HeapAlloc", "Ptr", this.hHeap, "UInt", Flags, "UPtr", Bytes, "UPtr"))
        Return this
    }

    /*
        Cambia la cantidad de memoria reservada. Antes de llamar a esta función debe reservar memoria utilizando Heap::Alloc.
        HEAP_NO_SERIALIZE            0x00000001
        HEAP_GENERATE_EXCEPTIONS     0x00000004
        HEAP_ZERO_MEMORY             0x00000008
        HEAP_REALLOC_IN_PLACE_ONLY   0x00000010
    */
    ReAlloc(Bytes, Flags := 4)
    {
        ; HeapReAlloc function
        ; https://msdn.microsoft.com/en-us/library/windows/desktop/aa366704(v=vs.85).aspx
        ObjRawSet(this, "Ptr", DllCall("Kernel32.dll\HeapReAlloc", "Ptr", this.hHeap, "UInt", Flags, "UPtr", this.Ptr, "UPtr", Bytes, "UPtr"))
        Return this
    }

    /*
        Libera toda la memoria reservada. Para volver a reservar memoria, debe utilizar Heap::Alloc.
        HEAP_NO_SERIALIZE   0x00000001
    */
    Free(Flags := 0)
    {
        ; HeapFree function
        ; https://msdn.microsoft.com/en-us/library/windows/desktop/aa366701(v=vs.85).aspx
        DllCall("Kernel32.dll\HeapFree", "Ptr", this.hHeap, "UInt", Flags, "UPtr", this.Ptr, "UInt")
        ObjRawSet(this, "Ptr", 0)
        Return this
    }

    /*
        Copia datos de la dirección de memoria especificada.
    */
    CopyFrom(Address, Offset, Bytes)
    {
        DllCall("msvcrt.dll\memcpy_s", "UPtr", this.Ptr+Offset, "UPtr", this.Size-Offset, "UPtr", Address, "UPtr", Bytes, "Cdecl")
        Return this
    }

    /*
        Copia datos en la dirección de memoria especificada.
    */
    CopyTo(Address, Offset, Bytes)
    {
        DllCall("msvcrt.dll\memcpy", "UPtr", Address, "UPtr", this.Ptr+Offset, "UPtr", Bytes, "Cdecl")
        Return this
    }

    /*
        Rellena la memoria con el valor especificado.
    */
    Fill(UChar, Offset := 0, Bytes := -1)    ; UChar   0-255
    {
        DllCall("NtDll.dll\RtlFillMemory", "UPtr", this.Ptr+Offset, "UPtr", Bytes == -1 ? this.Size-Offset : Bytes, "UChar", UChar)
        Return this
    }

    /*
        Crea una copia de la memoria actual y devuelve un nuevo objeto Heap.
    */
    Clone(Offset := 0, Bytes := -1)
    {
        Local Mem := new Heap().Alloc(Bytes == -1 ? this.Size-Offset : Bytes)
        Return this.CopyTo(Mem.Ptr, Offset, Mem.Size) ? Mem : 0
    }

    /*
        Compara con los datos en el búfer especificado.
    */
    Compare(Address, Offset := 0, Bytes := -1)
    {
        ; https://msdn.microsoft.com/es-es/library/zyaebf12.aspx
        Return DllCall("msvcrt.dll\memcmp", "UPtr", this.Ptr+Offset, "UPtr", Address, "UPtr", Bytes == -1 ? this.Size-Offset : Bytes, "CDecl")
    }

    /*
        Recorta los datos actuales.
    */
    Trim(Offset, Bytes := -1)
    {
        ; https://msdn.microsoft.com/es-es/library/8k35d1fx.aspx
        ; a diferencia de memcpy, memmove asegura que los bytes originales sean copiados correctamente si se superponen
        DllCall("msvcrt.dll\memmove", "UPtr", this.Ptr, "UPtr", this.Ptr+Offset, "UPtr", Bytes := Bytes == -1 ? this.Size-Offset : Bytes, "Cdecl")
        Return this.ReAlloc(Bytes)
    }

    /*
        Busca caracteres en el búfer.
    */
    Chr(UChar, Offset := 0, Count := -1)
    {
        ; https://msdn.microsoft.com/es-es/library/d7zdhf37.aspx
        Return DllCall("msvcrt.dll\memchr", "UPtr", this.Ptr+Offset, "UChar", UChar, "UPtr", Count == -1 ? this.Size-Offset : Count, "CDecl UPtr")
    }

    /*
        Libera toda la memoria reservada sobrante para la cadena actual.
    */
    StrTrim(u8 := 0)
    {
        Return this.ReAlloc((this.Length[u8] + 1) * (u8 ? 1 : 2))
    }

    /*
        Recupera o establece la cantidad de memoria actualmente reservada, en bytes.
    */
    Size[Flags := 4]
    {
        ; HEAP_NO_SERIALIZE   0x00000001
        Get
        {
            ; HeapSize function
            ; https://msdn.microsoft.com/en-us/library/windows/desktop/aa366706(v=vs.85).aspx
            Return DllCall("Kernel32.dll\HeapSize", "Ptr", this.hHeap, "UInt", !!Flags, "UPtr", this.Ptr, "UPtr")
        }

        Set
        {
            ObjRawSet(this, "Ptr", DllCall("Kernel32.dll\HeapReAlloc", "Ptr", this.hHeap, "UInt", Flags, "UPtr", this.Ptr, "UPtr", Value, "UPtr"))
            Return Value
        }
    }

    /*
        Recupera la cantidad de caracteres o establece la longitud de la cadena.
    */
    Length[u8 := 0]
    {
        Get
        {
            ; https://msdn.microsoft.com/es-ar/library/z50ty2zh.aspx
            Return DllCall("msvcrt.dll\" . (u8?"strnlen":"wcsnlen"), "UPtr", this.Ptr, "UPtr", this.Size, "CDecl UPtr")
        }

        Set
        {
            NumPut(0, this.Ptr + Value * (u8 ? 1 : 2), u8 ? "UChar" : "UShort")
            Return Value
        }
    }
}
