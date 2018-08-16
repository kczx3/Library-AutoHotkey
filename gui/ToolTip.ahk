/*
    Crea ToolTip"s personalizados. Permite modificar el texto, título, ícono y fuente.
    Observaciones:
        Esta clase solo funciona a partir de Windows Vista en adelante.
        Ver ejemplo al final de la clase.
*/
Class _ToolTip    ; WIN_V+
{
    ; ===================================================================================================================
    ; INSTANCE VARIABLES
    ; ===================================================================================================================
    hWnd        := 0                        ; El identificador de la ventana ToolTip
    TOOLINFO    := ""                       ; La estructura TOOLINFO
    pTOOLINFO   := 0                        ; Puntero a la estructura TOOLINFO
    X           := 0                        ; Coordenada X. Esta variable se actualiza al usar la función Show cuando no se espesifica el punto X
    Y           := 0                        ; Coordenada Y. Esta variable se actualiza al usar la función Show cuando no se espesifica el punto Y
    W           := 0                        ; Ancho del ToolTip. Esta variable se actualiza al usar la función Show cuando no se espesifica el punto X
    H           := 0                        ; Alto del ToolTip. Esta variable se actualiza al usar la función Show cuando no se espesifica el punto Y
    
    
    ; ===================================================================================================================
    ; CONSTRUCTOR
    ; ===================================================================================================================
    /*
        Parámetros:
            Title: El título a mostrar para este ToolTip. Si este parámetro es una cadena vacía, el título no es mostrado.
            Text : El texto o mensaje a mostrar en este ToolTip. Si asigna una cadena vacía a un ToolTip, éste no podrá ser mostrado.
            Icon : Un valor que identifica al icono a mostrar. Este parámetro puede ser un identificador a un icono. Si este valor es 0, el icono no es mostrado.
                1 / 4 = Icono de información pequeño/grande.
                2 / 5 = Icono de advertencia pequeño/grande.
                3 / 6 = Icono de error pequeño/grande.
    */
    __New(Title := "", Text := "", Icon := 0)
    {
        ; https://msdn.microsoft.com/en-us/library/windows/desktop/ms632680(v=vs.85).aspx
        This.hWnd     := DllCall("User32.dll\CreateWindowExW", "UInt", 0x00000008         ;dwExStyle    --> WS_EX_TOPMOST
                                                             , "Str" , "tooltips_class32" ;dwExStyle    --> (https://msdn.microsoft.com/en-us/library/windows/desktop/bb760250(v=vs.85).aspx)
                                                             , "Ptr" , 0                  ;lpWindowName --> NULL
                                                             , "UInt", 0x80000003         ;dwStyle      --> WS_POPUP
                                                             , "Int" , 0x80000000         ;x            --> CW_USEDEFAULT
                                                             , "Int" , 0x80000000         ;y            --> CW_USEDEFAULT
                                                             , "Int" , 0x80000000         ;nWidth       --> CW_USEDEFAULT
                                                             , "Int" , 0x80000000         ;nHeight      --> CW_USEDEFAULT
                                                             , "Ptr" , 0                  ;hWndParent   --> NULL (ignored)
                                                             , "Ptr" , 0                  ;hMenu        --> NULL
                                                             , "Ptr" , 0                  ;hInstance    --> NULL
                                                             , "Ptr" , 0                  ;lpParam      --> NULL
                                                             , "Ptr")                     ;ReturnType
        Text .= ""    ; nos aseguramos que Text sea una cadena

        ObjSetCapacity(this, "TOOLINFO", 6*4 + 6*A_PtrSize)    ; reservamos memoria para la estructura TOOLINFO
        this.pTOOLINFO := ObjGetAddress(this, "TOOLINFO")    ; recuperamos la dirección de memoria de TOOLINFO

        NumPut(6*4 + 6*A_PtrSize       , this.pTOOLINFO                   , "UInt")  ;cbSize      --> tamaño de TOOLINFO
        NumPut(0x0080 | 0x0001 | 0x0020, this.pTOOLINFO + 4               , "UInt")  ;uFlags      --> opciones
        NumPut(This.hWnd               , this.pTOOLINFO + 2*4             , "Ptr")   ;hwnd        --> el identificador de la ventana ToolTip
        NumPut(This.hWnd               , this.pTOOLINFO + 2*4 + A_PtrSize , "Ptr")   ;uId         --> el identificador de la ventana propietaria del ToolTip
        NumPut(&Text                   , this.pTOOLINFO + 24 + 3*A_PtrSize)          ;lpszText    --> la dirección de memoria a una cadena que especifica el texto a mostrar

        ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb760338(v=vs.85).aspx
        DllCall("User32.dll\SendMessageW", "Ptr" , This.hWnd          ;Window handle
                                         , "UInt", 0x0432             ;TTM_ADDTOOLW
                                         , "Ptr" , 0                  ;wParam (Must be zero)
                                         , "UPtr", this.pTOOLINFO)    ;lParam (Pointer to a TOOLINFO structure)

        ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb760408(v=vs.85).aspx
        DllCall("User32.dll\SendMessageW", "Ptr" , This.hWnd    ;Window handle
                                         , "UInt", 0x0418       ;TTM_SETMAXTIPWIDTH
                                         , "Ptr" , 0            ;wParam (Must be zero)
                                         , "Int" , 0)           ;lParam (Maximum tooltip window width, or 0 to allow any width)

        This.SetTitle(Title, Icon)    ; establecemos el título
    }


    ; ===================================================================================================================
    ; DESTRUCTOR
    ; ===================================================================================================================
    __Delete()
    {
        Local hFont
        If (hFont := DllCall("User32.dll\SendMessageW", "Ptr", This.hWnd, "UInt", 0x0031, "Ptr", 0, "Ptr", 0, "Ptr"))    ; comprobamos si el usuario cambió la fuente por defecto del ToolTip
            DllCall("Gdi32.dll\DeleteObject", "Ptr", hFont)    ; si es así, debemos eliminarla
        
        DllCall("User32.dll\DestroyWindow", "Ptr", This.hWnd)    ; eliminamos la ventana ToolTip
    } ;https://msdn.microsoft.com/en-us/library/windows/desktop/ms632682(v=vs.85).aspx
    
    
    ; ===================================================================================================================
    ; PUBLIC METHODS
    ; ===================================================================================================================
    /*
        Muestra, oculta, cambia el texto, y/o cambia la posición de este ToolTip.
        Parámeros:
            Text : El nuevo texto a mostrar. Si este parámetro es una cadena vacía el ToolTip se oculta. El texto no es modificado si es el mismo.
            X / Y: Las nuevas coordenadas. Si especifica una cadena vacía, se utilizan las coordenadas del cursor y se ajustan para que el ToolTip sea siempre visible en pantalla.
    */
    Show(Text, X := "", Y := "")
    {
        Local Pos, X2, Y2, VW

        If (Text == "")    ; si el usuario no especifico ningún texto, ocultamos el ToolTip
            This.Hide()
        Else
        {
            If (!(This.Text == Text))    ; solo cambiamos el texto si el nuevo texto es diferente al actual; esto para evitar parpadeos molestos cuando el texto es largo
                This.Text := Text

            If (X == "" || Y == "")    ; si no se especifica alguna coordenada, utilizamos las coordenadas del cursor y las adaptamos para que el ToolTip se visualize correctamente en pantalla
            {
                Pos := This.GetPos()    ; recuperamos las dimensiones del ToolTip
                CoordMode("Mouse", "Screen")    ; nos aseguramos que MouseGetPos recupere las coordenadas en relación al escritorio (pantalla completa) y no a otra ventana
                MouseGetPos(X2, Y2)    ; recuperamos las coordenadas actuales del cursor
                
                If (X == "")
                {
                    X := X2 + 10
                    This.X := X := X + (This.W:=Pos.W) > (VW:=SysGet(78)) ? X - (X + Pos.W - VW) : X
                }

                If (Y == "")
                {
                    Y := Y2 + 10
                    This.Y := Y := Y + (This.H:=Pos.H) > SysGet(79) ? Y - Pos.H - 10 : Y
                }
            } 
            
            This.Move(X, Y)    ; re-posicionamos el ToolTip en las nuevas coordenadas
            If (!This.Visible)    ; si el ToolTip no es visible, lo hacemos visible
                This.Visible := TRUE
        }
    }
    
    /*
        Oculta este ToolTip.
    */
    Hide()
    {
        DllCall("User32.dll\SendMessageW", "Ptr", This.hWnd, "UInt", 0x0411, "Int", FALSE, "UPtr", this.pTOOLINFO)
    } ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb760421(v=vs.85).aspx
    
    /*
        Recupera las coordenadas y dimensiones de este ToolTip.
        Return:
            Devuelve un objeto con las claves X, Y, W y H.
    */
    GetPos()
    {
        Local X, Y
        WinGetPos(X, Y,,, "ahk_id" . This.hWnd)    ; recuperamos las coordenadas del ToolTip
        Local Size := DllCall("User32.dll\SendMessageW", "Ptr", This.hWnd, "UInt", 0x041E, "Ptr", 0, "UPtr", this.pTOOLINFO, "UInt")    ; recuperamos las dimensiones del ToolTip
        Return {X: X, Y: Y, W: Size & 0xFFFF, H: Size >> 16}
    } ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb760387(v=vs.85).aspx
    
    /*
        Mueve este ToolTip a las coordenadas especificadas.
        Parámetros:
            X / Y: Las nuevas coordenadas.
    */
    Move(X, Y)
    {
        DllCall("User32.dll\SendMessageW", "Ptr", This.hWnd, "UInt", 0x0412, "Ptr", 0, "UInt", (X & 0xFFFF) | (Y << 16))
    } ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb760422(v=vs.85).aspx
    
    /*
        Fuerza el redibujado de este ToolTip.
    */
    Update()
    {
        DllCall("User32.dll\SendMessageW", "Ptr", This.hWnd, "UInt", 0x0412, "Ptr", 0, "UInt", 0)
    } ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb760425(v=vs.85).aspx
    
    /*
        Cambia el título y el icono mostrado en este ToolTip.
        Parámetros:
            NewTitle: El nuevo título. Si este parámetro es una cadena vacía, el título es removido.
            Icon    : Un valor que identifica al icono a mostrar. Este parámetro puede ser un identificador a un icono. Si este valor es 0, el icono no es mostrado.
                1 / 4 = Icono de información pequeño/grande.
                2 / 5 = Icono de advertencia pequeño/grande.
                3 / 6 = Icono de error pequeño/grande. 
    */
    SetTitle(NewTitle, Icon := 0)
    {
        NewTitle .= ""    ; nos aseguramos de que NewTitle sea una cadena
        DllCall("User32.dll\SendMessageW", "Ptr", This.hWnd, "UInt", 0x0421, "Ptr", Icon, "UPtr", NewTitle == "" ? 0 : &NewTitle)
    } ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb760414(v=vs.85).aspx
    
    /*
        Cambia la fuente del texto de este ToolTip.
        Parámetros:
            Options: Las opciones de la fuente. Debe especificar una cadena con una o más de las siguientes palabras claves:
                sN                                 = El tamaño del texto. Por defecto es 9.
                qN                                 = La calidad de la fuente. Por defecto es 5 (ClearType).
                wN                                 = El peso del texto. 400 es normal, 600 es semi-negrita, 700 es negrita. Por defecto es 400.
                Italic / Underline / Strike / Bold = El estilo de la fuente. Cursiva / Subrayado / Tachado / Negrita.
            FontName: El nombre de la fuente. Si este parámetro es una cadena vacía, la fuente actual es removida y se reestablece a la fuente original.
        Nota:
            Si especifica el peso del texto (wN), "Bold" no tiene efecto; ya que "Bold" hace que wN sea 700 (negrita).
    */
    SetFont(Options := "", FontName := "Segoe UI")
    {
        Local hFont
        If (hFont := DllCall("User32.dll\SendMessageW", "Ptr", This.hWnd, "UInt", 0x0031, "Ptr", 0, "Ptr", 0, "Ptr"))    ; eliminamos la fuente asignada anteriormente, si la hay
            DllCall("Gdi32.dll\DeleteObject", "Ptr", hFont)

        If (FontName == "")    ; si FontName es una cadena vacía, restauramos la fuente original
            Return DllCall("User32.dll\SendMessageW", "Ptr", This.hWnd, "UInt", 0x0030, "Ptr", 0, "Int", TRUE)    ; WM_SETFONT = 0x0030

        Local hDC := DllCall("Gdi32.dll\CreateDCW", "Str", "DISPLAY", "Ptr", 0, "Ptr", 0, "Ptr", 0, "Ptr")
              , R := DllCall("Gdi32.dll\GetDeviceCaps", "Ptr", hDC, "Int", 90)
        DllCall("Gdi32.dll\DeleteDC", "Ptr", hDC)
            
        Local t
            , Size      := RegExMatch(Options, "i)s([\-\d\.]+)(p*)", t) ? t[1] : 10
            , Height    := Round((Abs(Size) * R) / 72) * -1
            , Quality   := RegExMatch(Options, "i)q([\-\d\.]+)(p*)", t) ? t[1] : 5
            , Weight    := RegExMatch(Options, "i)w([\-\d\.]+)(p*)", t) ? t[1] : (InStr(Options, "Bold") ? 700 : 400)
            , Italic    := !!InStr(Options, "Italic")
            , Underline := !!InStr(Options, "Underline")
            , Strike    := !!InStr(Options, "Strike")
            
        DllCall("User32.dll\SendMessageW", "Ptr", This.hWnd, "UInt", 0x0030
                                                           , "Ptr" , DllCall("Gdi32.dll\CreateFontW", "Int", Height, "Int", 0, "Int", 0, "Int", 0, "Int", Weight, "UInt", Italic, "UInt", Underline, "UInt", Strike, "UInt", 1, "UInt", 4, "UInt", 0, "UInt", Quality, "UInt", 0, "UPtr", &FontName, "Ptr")
                                                           , "Int" , TRUE)
    } ;https://msdn.microsoft.com/en-us/library/windows/desktop/ms632642(v=vs.85).aspx
    
    
    ; ===================================================================================================================
    ; PROPERTIES
    ; ===================================================================================================================
    /*
        Recupera o establece el texto mostrado en este ToolTip.
    */
    Text[]
    {
        Get
        {
            Local Buffer
            VarSetCapacity(Buffer, 5024 * 2)    ; reservamos memoria para almacenar 5024*2 bytes
            NumPut(&Buffer, this.pTOOLINFO + 24 + 3*A_PtrSize)    ; le pasamos la dirección de memoria de Buffer a la estructura TOOLINFO que será utilizada por SendMessage::0x0438

            DllCall("User32.dll\SendMessageW", "Ptr", This.hWnd, "UInt", 0x0438, "Ptr", 5024, "UPtr", this.pTOOLINFO)    ; este mensaje recupera el texto del ToolTip y lo escribe en Buffer
            Return StrGet(&Buffer, "UTF-16")    ; devolvemos la cadena
        }

        Set
        {
            Value .= ""    ; nos aseguramos que Value es una cadena
            NumPut(&Value, this.pTOOLINFO + 24 + 3*A_PtrSize)    ; le pasamos la dirección de memoria de Value a la estructura TOOLINFO que será utilizada por SendMessage::0x0439
            DllCall("User32.dll\SendMessageW", "Ptr", This.hWnd, "UInt", 0x0439, "Int", 0, "UPtr", this.pTOOLINFO)    ; actualizamos el texto del ToolTip (el contenido de Value es copiado)
        }
    }
    
    /*
        Determina o establece si este ToolTip es visible.
    */
    Visible[]
    {
        Get
        {
            Return DllCall("User32.dll\IsWindowVisible", "Ptr", This.hWnd)
        }
        Set
        {
            DllCall("User32.dll\SendMessageW", "Ptr", This.hWnd, "UInt", 0x0411, "Int", !!Value, "UPtr", this.pTOOLINFO)
        }
    }
}










; :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
; ::: EJEMPLO
; :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
/*
    TT := New _ToolTip()     ; creamos la ventana ToolTip TT
    TT.SetTitle("ToolTip 1", 1)    ; establecemos el título y el icono
    TT.SetFont("s12 Italic Underline", "Courier New")    ; establecemos la fuente, el tipo de fuente y el tamaño

    TT2 := New _ToolTip("ToolTip 2",, 2)    ; creamos una segunda ventana ToolTip TT2

    Loop    ; bucle infinito - presione la tecla ESCAPE para terminar el script
    {
        TT.Show("My ToolTip Text")    ; muestra la ventana ToolTip TT
        TT2.Show("ToolTip 1: x" . TT.X . " y" . TT.Y . " w" . TT.W . " h" . TT.H, 10, 10)    ; muestra la ventana ToolTip TT2
            
        Sleep(50)    ; esperamos 50 ms
    }
    Return

    ~Esc::ExitApp    ; ESCAPE para terminar
*/
