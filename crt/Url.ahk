/*
    Codifica una cadena en formato Url.
    Parámetros:
        Url     : La cadena de caracteres a codificar.
        Encoding: Codificación a usar. El estándar es UTF-8. UTF-16 es una implementación no estándar y no siempre es reconocida.
    Ejemplos:
        MsgBox("UTF-8:`n-----------------`nEncoded: " . (e:=URLEncode(t:="•ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_.~ÁÑñ")) . "`n`nDecoded: " . URLDecode(e) . "`n`nOriginal: " . t)
        MsgBox("UTF-16:`n-----------------`nEncoded: " . (e:=URLEncode(t:="•ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_.~ÁÑñ", "UTF-16")) . "`n`nDecoded: " . URLDecode(e) . "`n`nOriginal: " . t)
*/
URLEncode(Url, Encoding := "UTF-8")
{
    Static Unreserved := "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_.~"

    If (Url == "")
        return ""

    Local Encoded := ""
    If (Encoding = "UTF-16")
        Loop Parse, Url
            Encoded .= InStr(Unreserved, A_LoopField) ? A_LoopField : Format("%u{:04X}", Ord(A_LoopField))
    else if (Encoding = "UTF-8")
    {
        Local Buffer := ""
            ,   Code := 0x00
        VarSetCapacity(Buffer, StrPut(Url, "UTF-8")), StrPut(Url, &Buffer, "UTF-8")
        While (Code := NumGet(&Buffer + A_Index - 1, "UChar"))
            Encoded .= InStr(Unreserved, Chr(Code)) ? Chr(Code) : Format("%{:02X}", Code)
    }
    else
        Throw Exception("Function URLEncode Parameter #2 invalid",, SubStr(Encoding, 1, 50))

    Return Encoded
} ;http://rosettacode.org/wiki/URL_encoding#AutoHotkey | https://en.wikipedia.org/wiki/Percent-encoding





/*
    Decodifica una cadena en formato de Url (codificación de URL o por ciento).
    Parámetros:
        Url: La cadena de caracteres a decodificar. La codificación es detectada automáticamente.
*/
URLDecode(Url)
{
    Local R := "", T := 0
        , Encoding := InStr(Url, Chr(37) . "u") ? "UTF-16" : "UTF-8"
        , Trim     := Encoding == "UTF-16"      ? 2        : 1           ;%u     : %
        , Length   := Encoding == "UTF-16"      ? 4        : 2           ;0x0000 : 0x00
    
    Loop Parse, Url
        R .= A_LoopField == Chr(37) ? Chr("0x" . SubStr(Url, A_Index + Trim, T:=Length)) : (--T > -Trim ? "" : A_LoopField)
    
    If (Encoding == "UTF-8")
    {
        Local Buffer := ""
        VarSetCapacity(Buffer, StrPut(R, "UTF-8"))
        Loop Parse, R
            NumPut(Ord(A_LoopField), &Buffer + A_Index - 1, "UChar")
    }
    
    Return Encoding == "UTF-8" ? StrGet(&Buffer, "UTF-8") : R
} ;https://autohotkey.com/boards/viewtopic.php?t=4868
