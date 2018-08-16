/*    ---- EXAMPLE ----
Gui := GuiCreate()
TB1 := new Toolbar(Gui, "x5 y5 w550 h40 Border")
    TB1.OnEvent("Click", "Toolbar_Event_Click")
TB2 := new Toolbar(Gui, "x5 y50 w550 h" . SysGet(11) . " Border")
    TB2.OnEvent("Click", "Toolbar_Event_Click")
TB3 := new Toolbar(Gui, "x5 y87 w40 h360 Border Vertical")
    TB3.OnEvent("Click", "Toolbar3_Event_Click")

TB1.SetImageList(IL_Create())
Loop 10 + 10 + 10 + 10
    IL_Add(TB1.GetImageList(), "shell32.dll", -A_Index)

Loop 10
    TB1.Add(, "Button " . A_Index, A_Index-1,,, 1000+A_Index, 2000+A_Index)
TB1.SetButtonSize(55, 40)
TB1.AutoSize()

TB2.SetImageList(TB1.GetImageList())
Loop 10
    TB2.Add(, 0, 10+A_Index,,, 3000+A_Index, 4000+A_Index)
TB2.Add()
Loop 8
    TB2.Add(, 0, 20+A_Index,,, 5000+A_Index, 6000+A_Index)
TB2.SetButtonSize(SysGet(11), SysGet(11))
TB2.AutoSize()

TB3.SetImageList(TB1.GetImageList())
Loop 9
    TB3.Add(, 0, 30+A_Index, "Wrap",,, A_Index)    ; 4 = TBSTATE_ENABLED | 0x20 = TBSTATE_WRAP (for vertical ToolBars)
TB3.SetButtonSize(40, 40)
TB3.AutoSize()

Gui.Show("w560 h452")
    Gui.OnEvent("Close", "ExitApp")
return

Toolbar_Event_Click(TB, Identifier, Data, X, Y, IsRightClick)
{
    ToolTip "TB.Type " . TB.Type . "`nIdentifier " . Identifier . "`nData " . Data . "`n(X;Y) " . X . ";" . Y . "`nIsRightClick " . IsRightClick
    SetTimer("ToolTip", -1000)
}

Toolbar3_Event_Click(TB, Identifier, Data, X, Y, IsRightClick)
{
    TB.CheckButton(Identifier, -1)
}
*/






Class Toolbar
{
    ; ===================================================================================================================
    ; STATIC/CLASS VARIABLES
    ; ===================================================================================================================
    static CtrlList := {}    ; almacena una lista con todos los controles Toolbar {ControlID:ToolbarObj}


    ; ===================================================================================================================
    ; CONSTRUCTOR
    ; ===================================================================================================================
    /*
        Añade una barra de herramientas en la ventana GUI especificada.
        Parámetros:
            Gui:
                El objeto de ventana GUI. También puede especificar un objeto control existente (o su identificador).
            Options:
                Las opciones para el nuevo control.
    */
    __New(Gui, Options := "")
    {
        if (Type(Gui) != "Gui")
        {
            Gui := IsObject(Gui) ? Gui.Hwnd : Gui
            local hWnd := 0, obj := ""
            For hWnd, obj in Toolbar.CtrlList
                if (hWnd == Gui)
                    return obj
            return 0
        }

        local Style := 0x40000000 | 0x02000000 | 0x8 | 0x40 | 0x4
        Style |= (InStr(Options, "Menu")                            ? 0x800|0x1000|0x04000000 :0)
               | (InStr(Options, "Vertical")                        ? 0x00080                 :0)
               | (InStr(Options, "Wrapable")                        ? 0x00200                 :0)
               | (InStr(Options, "Tabstop")                         ? 0x10000                 :0)
               | (InStr(Options, "Nodivider")                       ? 0x00040                 :0)
               | (InStr(Options, "Adjustable")                      ? 0x00020                 :0)
               | (InStr(Options, "Tooltips")                        ? 0x00100                 :0)
               | (InStr(Options, "List") && !InStr(Options, "Menu") ? 0x01000                 :0)
               | (InStr(Options, "Flat")                            ? 0x00800                 :0)
               | (InStr(Options, "Bottom")                          ? 0x00003                 :0)
        local k := "", v := ""
        For k, v in ["Menu","Vertical","Wrapable","Tabstop","Nodivider","Adjustable","Tooltips","List","Flat","Bottom"]
            Options := RegExReplace(Options, "i)\b" . v . "\b")

        ; Toolbar Control Reference
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/bumper-toolbar-toolbar-control-reference
        this.ctrl := Gui.AddCustom("ClassToolbarWindow32 +0x" . Format("{:X}", Style) . A_Space . Options)
        this.hWnd := this.ctrl.Hwnd
        this.gui  := Gui
        this.Type := "Toolbar"

        this.Buffer := ""
        ObjSetCapacity(this, "Buffer", 48)
        this.ptr := ObjGetAddress(this, "Buffer")

        this.ExStyle := 8
        this.Callback := {NM_CLICK: 0}
        this.ctrl.OnNotify(-2, ObjBindMethod(this, "EventHandler", -2))
        this.ctrl.OnNotify(-5, ObjBindMethod(this, "EventHandler", -5))
        ObjRawSet(Toolbar.CtrlList, this.hWnd, this)

        ; TB_BUTTONSTRUCTSIZE message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-buttonstructsize
        DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x41E, "UInt", 8 + 3*A_PtrSize, "Ptr", 0)
    }


    ; ===================================================================================================================
    ; PRIVATE METHODS
    ; ===================================================================================================================
    EventHandler(NotifyCode, GuiControl, lParam)
    {
        local ret := 0    ; FALSE = allow default processing of the click
        if (NotifyCode == -2 || NotifyCode == -5)    ; -2 = NM_CLICK | -5 = NM_RCLICK
        {
            if (this.Callback.NM_CLICK)
            {
                ret := this.Callback.NM_CLICK.Call( this
                                                  , NumGet(lParam+3*A_PtrSize  , "Ptr" )     ; NMMOUSE.dwItemSpec
                                                  , NumGet(lParam+4*A_PtrSize  , "UPtr")     ; NMMOUSE.dwItemData
                                                  , NumGet(lParam+5*A_PtrSize  , "Int" )     ; NMMOUSE.pt.x
                                                  , NumGet(lParam+5*A_PtrSize+4, "Int" )     ; NMMOUSE.pt.y
                                                  , NotifyCode == -5 )                       ; IsRightClick
            }

        }
        return ret is "integer" ? ret : 0
    }

    _State(ByRef State)
    {
        return State is "integer" ? State : ( (InStr(State, "Checked")       ? 0x01 : 0x00)
                                            | (InStr(State, "Disabled")      ? 0x00 : 0x04)
                                            | (InStr(State, "Hidden")        ? 0x08 : 0x00)
                                            | (InStr(State, "Indeterminate") ? 0x10 : 0x00)
                                            | (InStr(State, "Pressed")       ? 0x02 : 0x00)
                                            | (InStr(State, "Wrap")          ? 0x20 : 0x00)
                                            | (InStr(State, "Marked")        ? 0x80 : 0x00)
                                            | (InStr(State, "Elipses")       ? 0x40 : 0x00) )
    }

    ; ===================================================================================================================
    ; PUBLIC METHODS
    ; ===================================================================================================================
    /*
        Elimina el control.
    */
    Destroy()
    {
        ObjDelete(Toolbar.CtrlList, this.hWnd)
      , DllCall("User32.dll\DestroyWindow", "Ptr", this.hWnd)
    }

    /*
        Añade un elemento en la posición especificada.
        Parámetros:
            Item:
                El índice basado en cero del nuevo elemento. Para insertar un elemento al final de la lista, establezca el parámetro en -1.
            Text:
                El texto del nuevo elemento. Puede ser una cadena vacía o una dirección de memoria si es un número de tipo entero (integer).
                Especifique el número cero de tipo entero (integer) para dejar el botón sin una acadena asignada. Cuando llame al método GetTextLength devolverá -1.
                Si establece este parámetro en el caracter TAB, añade un separador. Este es el valor por defecto.
            Image:
                El índice basado en cero de una imagen dentro de la lista de imágenes. El valor -2 indica que el botón no debe mostrar ninguna imagen (este es el valor por defecto).
                Si se va a añadir un separador, este parámetro determina el ancho del separador, en píxeles. Por defecto es 5 píxeles (valor -2).
            State:
                Establece los estado para el botón. Referencia: "https://docs.microsoft.com/es-es/windows/desktop/Controls/toolbar-button-states".
                Este parámetro es ignorado si se va a añadir un separador.
                Puede especificar una o más de las siguientes palabras o un número entero que define los estados.
                Checked        = El botón se mantiene resaltado.
                Disabled       = El botón esta deshabilitado.
                Hidden         = El botón esta oculto.
                Indeterminate  = El texto del botón esta siempre en gris y cuando se posiciona el cursor por encima del boton no es resaltado.
                Pressed        = El botón esta inicialmente precionado. un clic hará que vuelva a su estado normal.
                Wrap           = El botón es seguido por un salto de línea. Este estilo es necesario para Toolbars con botones en vertical.
                Marked         = El botón está marcado. La interpretación de un elemento marcado depende de la aplicación.
                Ellipses       = El texto del botón se corta y se muestra una elipsis.
            Style:
                Define el estilo del botón. Referencia: "https://docs.microsoft.com/es-es/windows/desktop/Controls/toolbar-control-and-button-styles".
            Data:
                Un número entero sin signo definido por el usuario para el botón. Útil para asignar datos a un botón, pasando una dirección de memoria.
                Este número es de 4 bytes en AHK de 32-bit y de 8 bytes en AHK de 64-bit.
            Identifier:
                Identificador de comando asociado con el botón. Este identificador se usa en un mensaje WM_COMMAND cuando se presiona el botón. Debe ser un número entero de 4 bytes.
                Este parámetro no es apto para pasar direcciones de memoria. Utilize el parámetro «Data» para asociar datos al botón.
        Return:
            Si tuvo éxito devuelve un valor distinto de cero.
    */
    Add(Item := -1, Text := "`t", Image := -2, State := 4, Style := 0, Data := 0, Identifier := 0)
    {
        ; _TBBUTTON structure
        ; https://docs.microsoft.com/en-us/windows/desktop/api/Commctrl/ns-commctrl-_tbbutton
        NumPut(Text == "`t" ? (Image == -2 ? 5 : Image) : Image, this.ptr, "Int")
      , NumPut(Identifier, this.ptr + 4, "Int")
      , NumPut(Text == "`t" ? 0x04 : this._State(State), this.ptr + 8, "UCHar")
      , NumPut(Text == "`t" ? 0x01 : Style, this.ptr + 9, "UCHar")
      , NumPut(Data, this.ptr + 8 + A_PtrSize, "UPtr")
      , NumPut(Type(Text) == "Integer" ? Text : Text == "`t" ? 0 : &Text, this.ptr + 8 + 2 * A_PtrSize, "UPtr")

        ; TB_INSERTBUTTON message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-insertbutton
        return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x443, "Ptr", Item, "UPtr", this.ptr)
    }
    
    /*
        Recupera información en el botón especificado.
        Return:
            Si tuvo éxito devuelve un objeto con las claves: Image, ID, State, Style, Data y Text. En caso contrario devuelve cero.
    */
    GetButton(Item)
    {
        ; TB_GETBUTTON message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-getbutton
        if (!DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x417, "Ptr", Item, "Ptr", this.ptr))
            return FALSE
        return { Image: NumGet(this.ptr            , "Int"  )
               ,    ID: NumGet(this.ptr+4          , "Int"  ) 
               , State: NumGet(this.ptr+8          , "UChar")
               , Style: NumGet(this.ptr+9          , "UChar")
               ,  Data: NumGet(this.ptr+8+A_PtrSize, "UPtr" )
               ,  Text: StrGet(NumGet(this.ptr+8+2*A_PtrSize, "UPtr")||(&""), "UTF-16") }
    }

    /*
        Recupera el texto de visualización en el botón especificado.
        Parámetros:
            Length:
                La cantidad máxima de caracteres a recuperar. Si este parámetro es -1, recupera el texto entero.
        Return:
            Si tuvo éxito devuelve el texto en el botón. En caso contrario devuelve una cadena vacía.
            ErrorLevel se establece en un valor distinto de cero si hubo un error, o cero en caso contrario.
            Tenga en cuenta que si ErrorLevel no es cero, no necesariamente quiere decir que el mensaje falló, puede ser que el botón no tenga una cadena asignada.
    */
    GetText(Identifier, Length := -1)
    {
        ; TB_GETBUTTONTEXT message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-getbuttontext
        local len := DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x44B, "Ptr", Identifier, "Ptr", 0, "Ptr")
        if ((len == -1 && (ErrorLevel := TRUE)) || !len)
            return ""
        local buffer := ""
        VarSetCapacity(buffer, len * 2 + 2)
      , ErrorLevel := !DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x44B, "Ptr", Identifier, "UPtr", &buffer, "Ptr")
        return ErrorLevel ? "" : Length == -1 ? StrGet(&buffer, len, "UTF-16") : SubStr(StrGet(&buffer, len, "UTF-16"), 1, Length)
    }

    /*
        Recupera la cantidad de caracteres en el texto asignado al botón especificado.
        Return:
            Devuelve la cantidad de caracteres en el texto. Si hubo un error o el botón no tiene una cadena asignada, devuelve -1.
    */
    GetTextLength(Identifier)
    {
        ; TB_GETBUTTONTEXT message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-getbuttontext
        return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x44B, "Ptr", Identifier, "Ptr", 0, "Ptr")
    }

    /*
        Establece la información para un botón existente en la barra de herramientas.
        Parámetros:
            Ver el método Toolbar::Add para la información de los parámetros.
            Command es el nuevo identificador (Identifier) para el botón.
            State debe ser un número entero y no una cadena (como otra opción) como en el método Toolbar:Add.
            Width es el ancho del botón.
            Los valores por defecto de los parámetros indican que no debe modificarse el valor actual.
        Return:
            Devuelve distinto de cero si tiene éxito, o cero de lo contrario.
    */
    SetButton(Identifier, Command := "", Image := -3, State := "", Style := "", Width := -1, Data := "", Text := -1)
    {
        ; TB_SETBUTTONINFO message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-setbuttoninfo
        NumPut(A_PtrSize == 4 ? 32 : 48, this.ptr, "UInt")
      , NumPut((Type(Text)=="Integer"&&Text<0?0:2) | (Command==""?0:0x20) | (Image<-2?0:1) | (State==""?0:4) | (Style==""?0:8) | (Width<0?0:0x40) | (Data==""?0:0x10), this.ptr+4, "UInt")
      , NumPut(Command == "" ? 0 : Command, this.ptr+8, "Int")
      , NumPut(Image, this.ptr+12, "Int")
      , NumPut(State == "" ? 0 : State, this.ptr+16, "UChar")
      , NumPut(Style == "" ? 0 : Style, this.ptr+17, "UChar")
      , NumPut(Width, this.ptr+18, "UShort")
      , NumPut(Data == "" ? 0 : Data, this.ptr+16+A_PtrSize, "UChar")
      , NumPut(Type(Text) == "Integer" ? Text : &Text, this.ptr+16+2*A_PtrSize, "UPtr")
        return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x440, "Ptr", Identifier, "UPtr", this.ptr)
    }

    /*
        Establece el texto de visualización en el botón especificado.
        Parámetros:
            Text:
                El texto para el botón. Puede ser una cadena vacía o una dirección de memoria si es un número de tipo entero (integer).
                Especifique el número cero de tipo entero (integer) para dejar el botón sin una acadena asignada. Cuando llame al método GetTextLength devolverá -1.
        Return:
            Devuelve distinto de cero si tiene éxito, o cero de lo contrario.
    */
    SetText(Identifier, Text := "")
    {
        ; TB_SETBUTTONINFO message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-setbuttoninfo
        NumPut(((A_PtrSize==4?32:48) & 0xFFFFFFFF) | ((2 & 0xFFFFFFFF) << 32), this.ptr, "UInt64")    ; 2 = TBIF_TEXT
      , NumPut(Type(Text) == "Integer" ? Text : &Text, this.ptr+16+2*A_PtrSize, "UPtr")
        return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x440, "Ptr", Identifier, "UPtr", this.ptr)
    }

    /*
        Recupera el ancho y la altura actuales de los botones de la barra de herramientas, en píxeles.
        Return:
            Devuelve un objeto con las claves «W» y «H».
    */
    GetButtonSize()
    {
        ; TB_GETBUTTONSIZE message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-getbuttonsize
        local size := DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x43A, "Ptr", 0, "Ptr", 0, "UInt")
        return {W: size & 0xFFFF, H: (size >> 16) & 0xFFFF}
    }

    /*
        Establece el tamaño real de los botones en la barra de herramientas.
        Observaciones:
            TB_SETBUTTONSIZE generalmente se debe llamar después de agregar botones.
    */
    SetButtonSize(Width := "", Height := "")
    {
        local size := this.GetButtonSize()
        Width := Width == "" ? size.W : Width, Height := Height == "" ? size.H : Height
        ; TB_SETBUTTONSIZE message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-setbuttonsize
        return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x41F, "Ptr", 0, "UInt", (Width & 0xFFFF) | ((Height & 0xFFFF) << 16))
    }

    /*
        Establece el ancho mínimo y máximo de los botones en el control de barra de herramientas.
        Return:
            Devuelve distinto de cero si tiene éxito, o cero de lo contrario.
    */
    SetButtonWidth(Min, Max)
    {
        ; TB_SETBUTTONWIDTH message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-setbuttonwidth
        return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x43B, "Ptr", 0, "UInt", (Min & 0xFFFF) | ((Max & 0xFFFF) << 16))
    }

    /*
        Obtiene el tamaño ideal de la barra de herramientas.
        Return:
            Devuelve un objeto con las claves Width y Height. Si hubo un error el valor de las claves se establecen en cero.
    */
    GetIdealSize()
    {
        ; TB_GETIDEALSIZE message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-getidealsize
        local W := DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x463, "Ptr", 0, "UPtr", this.ptr)
        local H := DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x463, "Ptr", 1, "UPtr", this.ptr + 4)
        return {W: W?NumGet(this.ptr, "Int"):0, H: H?NumGet(this.ptr+8, "Int"):0}
    }

    /*
        Recupera el índice basado en cero del elemento activo en la barra de herramientas.
        Return:
            Devuelve el índice del elemento activo, o -1 si no hay ningún elemento activo establecido.
            Los controles de barra de herramientas que no tienen el estilo TBSTYLE_FLAT no tienen elementos activos.
    */
    GetHotItem()
    {
        ; TB_GETHOTITEM message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-gethotitem
        return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x447, "Ptr", 0, "Ptr", 0, "Ptr")
    }

    /*
        Recupera el rectángulo delimitador de un botón en la barra de herramientas.
        Return:
            Si tuvo éxito devuelve un objeto con las claves: L (left), T (top), R (right) y B (bottom).
    */
    GetItemRect(Item)
    {
        ; TB_GETITEMRECT message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-getitemrect
        local ret := DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x41D, "Ptr", Item, "UPtr", this.ptr)
        return ret ? {L: NumGet(this.ptr, "Int"), T: NumGet(this.ptr+4, "Int"), R: NumGet(this.ptr+8, "Int"), B: NumGet(this.ptr+12, "Int")} : 0
    }

    /*
        Recupera un puntero a la interfaz IDropTarget para el control de barra de herramientas.
        ErrorLevel se establece en un código de error HRESULT.
        IDropTarget es utilizada por la barra de herramientas cuando los objetos se arrastran o se sueltan en ella.
    */
    GetObject()
    {
        ; TB_GETOBJECT message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-getobject
        DllCall("Ole32.dll\CLSIDFromString", "Str", "{00000122-0000-0000-C000-000000000046}", "UPtr", this.ptr)
        local IDropTarget := 0
        ErrorLevel := DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x43E, "UPtr", this.ptr, "PtrP", IDropTarget, "UInt")
        return IDropTarget
    }
    
    /*
        Elimina un botón de la barra de herramientas.
        Return:
            Devuelve un valor distinto de cero si tuvo éxito.
    */
    Delete(Item)
    {
        ; TB_DELETEBUTTON message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-deletebutton
        return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x416, "Ptr", Item, "Ptr", 0)
    }

    /*
        Elimina todos los botones de la barra de herramientas.
        Return:
            Devuelve el número de botones eliminados.
    */
    DeleteAll()
    {
        local i := this.GetCount()
        Loop i
            this.Delete(0)
        return i
    }

    /*
        Recupera el índice basado en cero para el botón asociado con el identificador de comando especificado.
        Return:
            Devuelve el índice basado en cero para el botón o -1 si el identificador de comando especificado no es válido.
    */
    CommandToIndex(Identifier)
    {
        ; TB_COMMANDTOINDEX message
        ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb787305(v=vs.85).aspx
        return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x419, "Ptr", Identifier, "Ptr", 0)
    }

    /*
        Recupera el identificador de comando asociado con el índice basado en cero del botón especificado.
        Return:
            Devuelve el identificador de comando para el botón o una cadena vacía si el índice basado en cero especificado no es válido.
    */
    IndexToCommand(Item)
    {
        ; TB_GETBUTTON message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-getbutton
        return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x417, "Ptr", Item, "Ptr", this.ptr) ? NumGet(this.ptr+4, "Int") : ""
    }

    /*
        Recupera la cantidad de botones actualmente en la barra de herramientas.
    */
    GetCount()
    {
        ; TB_BUTTONCOUNT message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-buttoncount
        return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x418, "Ptr", 0, "Ptr", 0, "Ptr")
    }

    /*
        Marca o desmarca un botón dado en la barra de herramientas.
        Parámetros:
            Check:
                Indica si marcar o desmarcar el botón especificado. Especificar -1 para invertir el estado actual.
    */
    CheckButton(Identifier, Check := TRUE)
    {
        ; TB_CHECKBUTTON message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-checkbutton
        return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x402, "Ptr", Identifier, "Ptr", Check == -1 ? !this.IsButtonChecked(Identifier) : !!Check)
    }

    /*
        Determina si el botón especificado en la barra de herramientas está marcado.
        Return:
            Devuelve un valor distinto de cero si el botón está marcado, o cero de lo contrario.
    */
    IsButtonChecked(Identifier)
    {
        ; TB_ISBUTTONCHECKED message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-isbuttonchecked
        return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x40A, "Ptr", Identifier, "Ptr", 0)
    }

    /*
        Habilita o deshabilita el botón especificado en la barra de herramientas.
        Parámetros:
            Enable:
                Indica si habilitar o deshabilitar el botón especificado. Especificar -1 para invertir el estado actual.
    */
    EnableButton(Identifier, Enable := TRUE)
    {
        ; TB_ENABLEBUTTON message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-enablebutton
        return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x401, "Ptr", Identifier, "Ptr", Enable == -1 ? !this.IsButtonEnabled(Identifier) : !!Enable)
    }

    /*
        Determina si el botón especificado en la barra de herramientas está habilitado.
        Return:
            Devuelve un valor distinto de cero si el botón está habilitado, o cero de lo contrario.
    */
    IsButtonEnabled(Identifier)
    {
        ; TB_ISBUTTONENABLED message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-isbuttonenabled
        return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x409, "Ptr", Identifier, "Ptr", 0)
    }

    /*
        Muestra u oculta el botón especificado en la barra de herramientas.
        Parámetros:
            Show:
                Indica si mostrar u ocultar el botón especificado. Especificar -1 para invertir el estado actual.
    */
    ShowButton(Identifier, Show := TRUE)
    {
        ; TB_HIDEBUTTON message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-hidebutton
        return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x404, "Ptr", Identifier, "Ptr", Show == -1 ? !this.IsButtonVisible(Identifier) : !!Show)
    }

    /*
        Determina si el botón especificado en la barra de herramientas es visible.
        Return:
            Devuelve un valor distinto de cero si el botón es visible, o cero de lo contrario.
    */
    IsButtonVisible(Identifier)
    {
        ; TB_ISBUTTONHIDDEN message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-isbuttonhidden
        return !DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x40C, "Ptr", Identifier, "Ptr", 0)
    }

    /*
        Causa el cambio de tamaño de la barra de herramientas.
        Observaciones:
            Una aplicación envía el mensaje TB_AUTOSIZE después de hacer que el tamaño de la barra de herramientas cambie al establecer el tamaño del botón o del mapa de bits o al agregar cadenas por primera vez.
        Return:
            Devuelve siempre el objeto Toolbar.
    */
    AutoSize()
    {
        ; TB_AUTOSIZE message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-autosize
        DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x421, "Ptr", 0, "Ptr", 0)
    }

    /*
        Muestra el cuadro de diálogo Personalizar Barra de Herramientas.
    */
    Customize()
    {
        ; TB_CUSTOMIZE message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-customize
        DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x41B, "Ptr", 0, "Ptr", 0)
    }

    /*
        Vuelve a dibujar el área ocupada por el control.
    */
    Redraw()
    {
        ; InvalidateRect function
        ; https://docs.microsoft.com/es-es/windows/desktop/api/winuser/nf-winuser-invalidaterect
        return DllCall("User32.dll\InvalidateRect", "Ptr", this.hWnd, "UPtr", 0, "Int", TRUE)
    }

    /*
        Establece el foco del teclado en el control.
    */
    Focus()
    {
        this.ctrl.Focus()
    }

    /*
        Cambia la fuente.
    */
    SetFont(Options, FontName := "")
    {
        this.ctrl.SetFont(Options, FontName)
    }

    /*
        Mueve y/o cambia el tamaño del control, opcionalmente lo vuelve a dibujar.
    */
    Move(Pos, Draw := FALSE)
    {
        this.ctrl.Move(Pos, Draw)
    }

    /*
        Registra una función para ser llamada cuando ocurre un evento.
        Parámetros:
            EventName:
                El tipo de evento, debe ser una de las cadenas especificadas a continuación.
                Click  = Cuando el usuario presiona el clic izquierdo o derecho en cualquier parte del control.
                         La función a llamar recibe los parámetros: Callback(CtrlObj,Identifier,Data,X,Y,IsRightClick).
            Callback:
                El nombre o referencia a una función a llamar. Especificar una cadena vacía para eliminar la función del registro.
                A continuación se detallan los parámetros comunes entre todos los eventos que recibe la función.
                CtrlObj     = Recibe el objeto Toolbar en el que se produjo el mensaje.
                Identifier  = El identificador de comando del botón donde se produjo el clic.
                Data        = Un número entero sin signo asignado al botón.
                X / Y       = Las coordenadas del cursor relativas al área del control.
        Toolbar Control Notifications:
            https://docs.microsoft.com/es-es/windows/desktop/Controls/bumper-toolbar-control-reference-notifications
    */
    OnEvent(EventName, Callback)
    {
        If (EventName = "Click")
            this.Callback.NM_CLICK := Type(Callback) != "String" ? Callback : Callback == "" ? 0 : Func(Callback)
        else
            throw Exception("Class Toolbar::OnEvent invalid parameter #1", -1)
    }

    /*
        Establece la lista de imágenes que la barra de herramientas usa para mostrar los botones que están en su estado predeterminado.
        Parámetros:
            ImageList:
                El identificador de la lista de imagenes.
            Index:
                El índice de la lista.
        Return:
            Devuelve el identificador de la lista de imágenes previamente asociada con el control, o devuelve 0 (NULL) si no se estableció previamente una lista de imágenes.
            Si el parámetro Destroy es distinto de cero, devuelve un valor distinto de cero si tuvo éxito.
    */
    SetImageList(ImageList, Index := 0)
    {
        ; TB_SETIMAGELIST message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-setimagelist
        return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x430, "Ptr", Index, "Ptr", ImageList, "Ptr")
    }

    /*
        Establece la lista de imágenes que la barra de herramientas usa para mostrar los botones que están en estado presionado.
    */
    SetPressedImageList(ImageList, Index := 0)
    {
        ; TB_SETPRESSEDIMAGELIST message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-setpressedimagelist
        return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x468, "Ptr", Index, "Ptr", ImageList, "Ptr")
    }

    /*
        Establece la lista de imágenes que usará el control de la barra de herramientas para mostrar los botones de acceso directo.
    */
    SetHotImageList(ImageList)
    {
        ; TB_SETHOTIMAGELIST message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-sethotimagelist
        return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x434, "Ptr", 0, "Ptr", ImageList, "Ptr")
    }

    /*
        Establece la lista de imágenes que usará el control de la barra de herramientas para mostrar los botones desactivados.
    */
    SetDisabledImageList(ImageList)
    {
        ; TB_SETDISABLEDIMAGELIST message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-setdisabledimagelist
        return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x436, "Ptr", 0, "Ptr", ImageList, "Ptr")
    }

    /*
        Recupera la lista de imágenes que utiliza el control de barra de herramientas para mostrar los botones en su estado predeterminado.
        Un control de barra de herramientas utiliza esta lista de imágenes para mostrar los botones cuando no están activos o deshabilitados.
    */
    GetImageList()
    {
        ; TB_GETIMAGELIST message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-getimagelist
        return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x431, "Ptr", 0, "Ptr", 0, "Ptr")
    }

    /*
        Obtiene la lista de imágenes que utiliza el control de barra de herramientas para mostrar botones en estado presionado.
    */
    GetPressedImageList()
    {
        ; TB_GETPRESSEDIMAGELIST message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-getpressedimagelist
        return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x469, "Ptr", 0, "Ptr", 0, "Ptr")
    }

    /*
        Recupera la lista de imágenes que utiliza el control de barra de herramientas para mostrar botones de acceso rápido.
    */
    GetHotImageList()
    {
        ; TB_GETHOTIMAGELIST message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-gethotimagelist
        return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x435, "Ptr", 0, "Ptr", 0, "Ptr")
    }

    /*
        Recupera la lista de imágenes que utiliza el control de barra de herramientas para mostrar los botones inactivos.
    */
    GetDisabledImageList()
    {
        ; TB_GETDISABLEDIMAGELIST message
        ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-getdisabledimagelist
        return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x437, "Ptr", 0, "Ptr", 0, "Ptr")
    }


    ; ===================================================================================================================
    ; PROPERTIES
    ; ===================================================================================================================
    /*
        Recupera o establece los estilos que están en uso en el control.
        set:
            Devuelve un valor que contiene los estilos previamente utilizados para el control.
    */
    Style[]
    {
        get {
            ; TB_GETSTYLE message
            ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb787350(v=vs.85).aspx
            return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x439, "Ptr", 0, "Ptr", 0, "UInt")
        }
        set {
            ; TB_SETSTYLE message
            ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb787459(v=vs.85).aspx
            return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x438, "Ptr", 0, "UInt", value, "UInt")
        }
    }

    /*
        Recupera o establece los estilos extendidos que están en uso en el control.
        set:
            Devuelve un valor que contiene los estilos extendidos previamente utilizados para el control.
    */
    ExStyle[]
    {
        get {
            ; TB_GETEXTENDEDSTYLE message
            ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-getextendedstyle
            return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x455, "Ptr", 0, "Ptr", 0, "UInt")
        }
        set {
            ; TB_SETEXTENDEDSTYLE message
            ; https://docs.microsoft.com/es-es/windows/desktop/Controls/tb-setextendedstyle
            return DllCall("User32.dll\SendMessageW", "Ptr", this.hWnd, "UInt", 0x454, "Ptr", 0, "UInt", value, "UInt")
        }
    }

    /*
        Recupera la posición y dimensiones del control.
    */
    Pos[]
    {
        get {
            return this.ctrl.Pos
        }
    }

    /*
        Determina si el control tiene el foco del teclado.
    */
    Focused[]
    {
        get {
            return this.ctrl.Focused
        }
    }

    /*
        Recupera o establece el estado de visibilidad del control.
        get:
            Devuelve cero si la ventana no es visible, 1 en caso contrario.
        set:
            Si la ventana estaba visible anteriormente, el valor de retorno es distinto de cero.
            Si la ventana estaba previamente oculta, el valor de retorno es cero.
    */
    Visible[]
    {
        get {
            ; IsWindowVisible function
            ; https://msdn.microsoft.com/en-us/library/windows/desktop/ms633530(v=vs.85).aspx
            return DllCall("User32.dll\IsWindowVisible", "Ptr", this.hWnd)
        }
        set {
            ; ShowWindow function
            ; https://msdn.microsoft.com/en-us/library/windows/desktop/ms633548(v=vs.85).aspx
            return DllCall("User32.dll\ShowWindow", "Ptr", this.hWnd, "Int", Value ? 8 : 0)
        }
    }

    /*
        Recupera o establece el estado habilitado/deshabilitado del control.
        get:
            Si la ventana esta habilitada devuelve un valor distinto de cero, o cero en caso contrario.
        set:
            Si la ventana estaba deshabilitada, el valor de retorno es distinto de cero.
            Si la ventana estaba habilitada, el valor de retorno es cero.
    */
    Enabled[]
    {
        get {
            ; IsWindowEnabled function
            ; https://msdn.microsoft.com/en-us/library/windows/desktop/ms646303(v=vs.85).aspx
            return DllCall("User32.dll\IsWindowEnabled", "Ptr", this.hWnd)
        }
        set {
            ; EnableWindow function
            ; https://msdn.microsoft.com/en-us/library/windows/desktop/ms646291(v=vs.85).aspx
            return DllCall("User32.dll\EnableWindow", "Ptr", this.hWnd, "Int", !!Value)
        }
    }
}

ToolbarCreate(Gui, Options := "")
{
    return new Toolbar(Gui, Options)
}
