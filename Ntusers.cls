VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NTUsers"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

   Public Function AddUser(ByVal UserName As String, _
                           ByVal PWD As String) As Long
   Dim Result As Long, UNPtr As Long, PWDPtr As Long, ParmError As Long
   Dim SNArray() As Byte, UNArray() As Byte, PWDArray() As Byte
   Dim UserStruct As TUser1
   If Not EsOkUserName(UserName) Then
        AddUser = 1000
        Exit Function
   End If
   On Error GoTo Errorhandler
   '
   ' Move to byte arrays
   '
     SNArray = "" & vbNullChar
     UNArray = UserName & vbNullChar
     PWDArray = PWD & vbNullChar
   '
   ' Allocate buffer space
   '
     Result = NetAPIBufferAllocate(UBound(UNArray) + 1, UNPtr)
     Result = NetAPIBufferAllocate(UBound(PWDArray) + 1, PWDPtr)
   '
   ' Copy arrays to the buffer
   '
     Result = StrToPtr(UNPtr, UNArray(0))
     Result = StrToPtr(PWDPtr, PWDArray(0))
   '
   ' Fill the structure
   '
     With UserStruct
       .ptrName = UNPtr
       .ptrPassword = PWDPtr
       .dwPasswordAge = 3
       .dwPriv = USER_PRIV_USER
       .ptrHomeDir = 0
       .ptrComment = 0
       .dwFlags = UF_NORMAL_ACCOUNT Or UF_SCRIPT Or UF_PASSWD_CANT_CHANGE Or UF_DONT_EXPIRE_PASSWD
'      .dwFlags = &H10220
       .ptrScriptPath = 0
     End With
   '
   ' Add the user
   '
     Result = NetUserAdd1(SNArray(0), 1, UserStruct, ParmError)
     AddUser = Result
     'MsgBox ("res:" & Result & " , param :" & ParmError)
     Result = NetAPIBufferFree(UNPtr)
     Result = NetAPIBufferFree(PWDPtr)
    Exit Function
Errorhandler:
     AddUser = CLng(Err.Number)

   End Function
   Public Function AddMoreUsers(ByVal Ruta As String) As String
   '_______________________
   
   Const Cl = "Línea"
   Const P2 = ":"
   Dim Pos%, Ok% 'as Integer
   Dim Linea$   'as string
   Dim NumLinea% 'As Integer
   Dim UserName$ 'As string
   Dim PWD$     'As string
   Dim Res&     'As Long
   Dim t As Integer
   Dim ResFinal$ 'as string
   
   'Valores de retorno: _
        Exito           "" _
        Fracaso         0 bien "No puede abrirse el fichero de datos." _
                        o bien "El fichero de datos está vacío." _
                        o una lista especificando errores en lista: _
        Error en una linea               "Línea X: Nombre de usuario incorrecto." _
                                         "Línea X: El usuario ya existe." _
                                         "Línea X: No existe caracter de separación.
    ResFinal$ = EnCelda(ConFont(NKS("Linea", 1), 3, "Arial Black", "Navy"), 1, "#CCEADB", 15) & _
                EnCelda(ConFont(NKS("Contenido", 1), 3, "Arial Black", "Navy"), 2, "#CCEADB", 25) & _
                EnCelda(ConFont(NKS("Resultado", 1), 3, "Arial Black", "Navy"), 3, "#CCEADB", 60)
    On Error Resume Next
    Open Ruta For Input As #1
    If Err.Number > 0 Then
        ResFinal$ = EnTabla(EnCelda(ConFont("No puede abrirse el fichero de datos.", 2, "Arial Black", "Red"), 4, , , "CENTER"))
        AddMoreUsers = ResFinal$
        Exit Function
    End If
    
    On Error GoTo 0
    NumLinea% = 0: Ok% = 0
    Do While Not EOF(1)
        NumLinea% = NumLinea% + 1
        Line Input #1, Linea$
        Pos% = InStr(2, Linea$, P2)
        If Pos% = 0 Then
            If Linea$ = "" Then
                ResFinal$ = ResFinal$ & EnCelda(ConFont(Cl & Str(NumLinea%)), 1) & _
                                EnCelda(ConFont("&nbsp;"), 2) & _
                                EnCelda(ConFont("Linea vacía.", 2, "Arial Black", "Red"), 3)
            Else
                ResFinal$ = ResFinal$ & EnCelda(ConFont(Cl & Str(NumLinea%)), 1) & _
                                EnCelda(ConFont(Linea$), 2) & _
                                EnCelda(ConFont("No existe carácter de separación.", 2, "Arial Black", "Red"), 3)
            End If
        Else
            UserName$ = Left(Linea$, Pos% - 1)
            PWD$ = Right(Linea$, Len(Linea$) - Pos%)
            Res& = AddUser(UserName$, PWD$)
            Select Case Res&
               Case 0 'El usuario se ha creado porque no existía
                    ResFinal$ = ResFinal$ & EnCelda(ConFont(Cl & Str(NumLinea%)), 1) & _
                                EnCelda(ConFont(Linea$), 2) & _
                                EnCelda(ConFont("""" & UserName$ & _
                                """ se creó con éxito.", _
                                2, "Arial Black", "Green"), 3)
                    Ok% = Ok% + 1
                Case 1000
                    'Close #1
                    ResFinal$ = ResFinal$ & EnCelda(ConFont(Cl & Str(NumLinea%)), 1) & _
                                EnCelda(ConFont(Linea$), 2) & _
                                EnCelda(ConFont("""" & UserName$ & _
                                    """ no es un nombre de usuario correcto.", _
                                    2, "Arial Black", "Red"), 3)
                Case 2224 'El usuario existía
                    ResFinal$ = ResFinal$ & EnCelda(ConFont(Cl & Str(NumLinea%)), 1) & _
                                EnCelda(ConFont(Linea$), 2) & _
                                EnCelda(ConFont("""" & UserName$ & _
                                """ ya existía .", _
                                2, "Arial Black", "Red"), 3)
                Case 2245    'El password es demasiado corto, _
                            o demasiado largo, o caracteres únicos insuficientes.
                     ResFinal$ = ResFinal$ & EnCelda(ConFont(Cl & Str(NumLinea%)), 1) & _
                                EnCelda(ConFont(Linea$), 2) & _
                                EnCelda(ConFont("El password no es válido", _
                                2, "Arial Black", "Red"), 3)
               Case Else  'Error inesperado
                    ResFinal$ = ResFinal$ & EnCelda(ConFont(Cl & Str(NumLinea%)), 1) & _
                                EnCelda(ConFont(Linea$), 2) & _
                                EnCelda(ConFont("Error inesperado.""" & UserName$ & _
                                """ NO fue creado.", _
                                2, "Arial Black", "Red"), 3)
            End Select
        End If
    Loop
    Close #1
    If NumLinea% = 0 Then
        ResFinal$ = EnTabla(EnCelda(ConFont("El fichero de datos está vacío.", 2, "Arial Black", "Red"), 4, , , "CENTER"))
    Else
        ResFinal$ = EnTabla(ResFinal$) & "<BR>" & _
                    EnTabla(EnCelda(ConFont("Se crearon" & _
                    Str(Ok%) & " usuarios de" & _
                    Str(NumLinea%) & " requeridos.", 3, "Arial Black", "Navy"), 4, "#CCEADB", , "CENTER"))
    End If

   AddMoreUsers = ResFinal$
   
   End Function





Public Function AddMoreUsersToLocalGroup(ByVal Ruta As String, _
                                            ByVal Grupo As String) _
          As String
   '_______________________
   
   Const Cl = "Línea"
   Const P2 = ":"
   Dim Pos%, Ok% 'as Integer
   Dim Linea$   'as string
   Dim NumLinea% 'As Integer
   Dim UserName$ 'As string
   Dim PWD$     'As string
   Dim Res&     'As Long
   Dim t As Integer
   Dim ResFinal$ 'as string
   
   'Valores de retorno: _
        Exito           "" _
        Fracaso         0 bien "No puede abrirse el fichero de datos." _
                        o bien "El fichero de datos está vacío." _
                        o una lista especificando errores en lista: _
        Error en una linea               "Línea X: Nombre de usuario incorrecto." _
                                         "Línea X: El usuario ya existe." _
                                         "Línea X: No existe caracter de separación.
    ResFinal$ = EnCelda(ConFont(NKS("Linea", 1), 3, "Arial Black", "Navy"), 1, "#CCEADB", 15) & _
                EnCelda(ConFont(NKS("Contenido", 1), 3, "Arial Black", "Navy"), 2, "#CCEADB", 25) & _
                EnCelda(ConFont(NKS("Resultado", 1), 3, "Arial Black", "Navy"), 3, "#CCEADB", 60)
    On Error Resume Next
    Open Ruta For Input As #1
    If Err.Number > 0 Then
        ResFinal$ = EnTabla(EnCelda(ConFont("No puede abrirse el fichero de datos.", 2, "Arial Black", "Red"), 4, , , "CENTER"))
        AddMoreUsersToLocalGroup = ResFinal$
        Exit Function
    End If
    
    On Error GoTo 0
    NumLinea% = 0: Ok% = 0
    Do While Not EOF(1)
        NumLinea% = NumLinea% + 1
        Line Input #1, Linea$
        Pos% = InStr(2, Linea$, P2)
        If Pos% = 0 Then
            If Linea$ = "" Then
                ResFinal$ = ResFinal$ & EnCelda(ConFont(Cl & Str(NumLinea%)), 1) & _
                                EnCelda(ConFont("&nbsp;"), 2) & _
                                EnCelda(ConFont("Linea vacía.", 2, "Arial Black", "Red"), 3)
            Else
                ResFinal$ = ResFinal$ & EnCelda(ConFont(Cl & Str(NumLinea%)), 1) & _
                                EnCelda(ConFont(Linea$), 2) & _
                                EnCelda(ConFont("No existe carácter de separación.", 2, "Arial Black", "Red"), 3)
            End If
        Else
            UserName$ = Left(Linea$, Pos% - 1)
            PWD$ = Right(Linea$, Len(Linea$) - Pos%)
            Res& = AddUser(UserName$, PWD$)
            If Res& = 1000 Then
                'Close #1
                ResFinal$ = ResFinal$ & EnCelda(ConFont(Cl & Str(NumLinea%)), 1) & _
                            EnCelda(ConFont(Linea$), 2) & _
                            EnCelda(ConFont("""" & UserName$ & _
                                """ no es un nombre de usuario correcto.", _
                                2, "Arial Black", "Red"), 3)
            Else
                If Res& = 2224 Then 'El usuario existía
                    Res& = AddUserToLocalGroup(Grupo, UserName$)
                    If Res& = 1378 Then
                        ResFinal$ = ResFinal$ & EnCelda(ConFont(Cl & Str(NumLinea%)), 1) & _
                                    EnCelda(ConFont(Linea$), 2) & _
                                    EnCelda(ConFont("""" & UserName$ & _
                                    """ ya existía en el Grupo """ & _
                                    Grupo & """", _
                                    2, "Arial Black", "Red"), 3)
                    Else
                        If Res& = 2220 Then 'El grupo no existe
                            ResFinal$ = EnTabla(EnCelda(ConFont("El grupo local no existe. La operación es imposible.", 2, "Arial Black", "Red"), 4, , , "CENTER"))
                            AddMoreUsersToLocalGroup = ResFinal$
                            Exit Function
                       Else
                            ResFinal$ = ResFinal$ & EnCelda(ConFont(Cl & Str(NumLinea%)), 1) & _
                                        EnCelda(ConFont(Linea$), 2) & _
                                        EnCelda(ConFont("""" & UserName$ & _
                                        """ se añadió en el Grupo """ & _
                                        Grupo & """ con éxito.", _
                                        2, "Arial Black", "Green"), 3)
                            Ok% = Ok% + 1
                        End If
                    End If
                Else 'El usuario se ha creado porque no existía
                    Res& = AddUserToLocalGroup(Grupo, UserName$)
                    If Res& = 1378 Then 'El usuario ya estaba en el grupo.
                        ResFinal$ = ResFinal$ & EnCelda(ConFont(Cl & Str(NumLinea%)), 1) & _
                                    EnCelda(ConFont(Linea$), 2) & _
                                    EnCelda(ConFont("""" & UserName$ & _
                                    """ ya existía en el Grupo """ & _
                                    Grupo & """", _
                                    2, "Arial Black", "Red"), 3)
                    Else
                        If Res& = 2220 Then 'El grupo no existe
                            ResFinal$ = EnTabla(EnCelda(ConFont("El grupo local no existe. La operación es imposible.", 2, "Arial Black", "Red"), 4, , , "CENTER"))
                            Res& = DelUser(UserName$)
                            AddMoreUsersToLocalGroup = ResFinal$
                            Exit Function
                       Else
                            ResFinal$ = ResFinal$ & EnCelda(ConFont(Cl & Str(NumLinea%)), 1) & _
                                        EnCelda(ConFont(Linea$), 2) & _
                                        EnCelda(ConFont("""" & UserName$ & _
                                        """ se añadió en el Grupo """ & _
                                        Grupo & """ con éxito.", _
                                        2, "Arial Black", "Green"), 3)
                            Ok% = Ok% + 1
                        End If
                    End If
                End If
            End If
        End If
    Loop
    Close #1
    If NumLinea% = 0 Then
        ResFinal$ = EnTabla(EnCelda(ConFont("El fichero de datos está vacío.", 2, "Arial Black", "Red"), 4, , , "CENTER"))
    Else
        ResFinal$ = EnTabla(ResFinal$) & "<BR>" & _
                    EnTabla(EnCelda(ConFont("Se añadieron" & _
                    Str(Ok%) & " usuarios de" & _
                    Str(NumLinea%) & " requeridos.", 3, "Arial Black", "Navy"), 4, "#CCEADB", , "CENTER"))
    End If

   AddMoreUsersToLocalGroup = ResFinal$
   
   End Function

   Function AddUserToLocalGroup(ByVal GroupName As String, _
                                ByVal UserName As String) As Long
   
   'NetLocalGroupAddMembers ( _
    IN  LPCWSTR     servername OPTIONAL, _
    IN  LPCWSTR     groupname, _
    IN  DWORD      level, _
    IN  LPBYTE     buf, _
    IN  DWORD      totalentries
    'Valores de Retorno:        1387    El usuario no existe.
    '                           1378    El usuario ya estaba en ese grupo
    '
   Dim SNArray() As Byte, GNArray() As Byte, UNArray() As Byte
   Dim LocalGroup3Struct As LocalGroup3
   Dim Result As Long, UNPtr As Long
     
     SNArray = "" & vbNullChar
     GNArray = GroupName & vbNullChar
     UNArray = UserName & vbNullChar
     
    ' Allocate buffer space
   '
     Result = NetAPIBufferAllocate(UBound(UNArray) + 1, UNPtr)
     
   ' Copy arrays to the buffer
   '
     Result = StrToPtr(UNPtr, UNArray(0))

   LocalGroup3Struct.dominioynombre = UNPtr
  
     
     'Result = NetGroupAddUser(SNArray(0), GNArray(0), UNArray(0))
     Result = NetLocalGroupAddMembers(SNArray(0), GNArray(0), 3, LocalGroup3Struct, 1)
     
     AddUserToLocalGroup = Result
        
   End Function


      Public Function ChangePassword(ByVal Uname As String, _
                                  ByVal PWD As String) As Long
   
   Dim Result As Long, UNPtr As Long, PWDPtr As Long, ParmError As Long
   Dim SNArray() As Byte, UNArray() As Byte, PWDArray() As Byte
   Dim UserStruct As TUser1
   If Not EsOkUserName(Uname) Then
        ChangePassword = 1000
        Exit Function
   End If
   On Error GoTo Errorhandler
   '
   ' Move to byte arrays
   '
     SNArray = "" & vbNullChar
     UNArray = Uname & vbNullChar
     PWDArray = PWD & vbNullChar 'UnPass1
   '
   ' Allocate buffer space
   '
     Result = NetAPIBufferAllocate(UBound(UNArray) + 1, UNPtr)
     Result = NetAPIBufferAllocate(UBound(PWDArray) + 1, PWDPtr)
   '
   ' Copy arrays to the buffer
   '
     Result = StrToPtr(UNPtr, UNArray(0))
     Result = StrToPtr(PWDPtr, PWDArray(0))
   '
   ' Fill the structure
   '
     With UserStruct
       .ptrName = UNPtr
       .ptrPassword = PWDPtr
       '.dwPasswordAge = 3
       '.dwPriv = USER_PRIV_USER
       '.ptrHomeDir = 0
       '.ptrComment = 0
       .dwFlags = UF_NORMAL_ACCOUNT Or UF_SCRIPT Or UF_PASSWD_CANT_CHANGE Or UF_DONT_EXPIRE_PASSWD
'      '.dwFlags = &H10220
       '.ptrScriptPath = 0
     End With
     
     Result = NetUserSetInfo(SNArray(0), UNArray(0), 1, UserStruct, ParmError)
     ChangePassword = Result
     'MsgBox ("res:" & Result & " , param :" & ParmError)
     Result = NetAPIBufferFree(UNPtr)
     Result = NetAPIBufferFree(PWDPtr)
    Exit Function
Errorhandler:
     
     ChangePassword = CLng(Err.Number)

   End Function
Public Function RenameUser(ByVal OldUname As String, _
                              ByVal NewUname As String) As Long
   
   Dim Result As Long, UNPtr As Long, ParmError As Long
   Dim SNArray() As Byte, UNArray1() As Byte, UNArray2() As Byte
   Dim UserStruct As TUser0
   If Not EsOkUserName(NewUname) Then
        RenameUser = 1000
        Exit Function
   End If
   On Error GoTo Errorhandler
   '
   ' Move to byte arrays
   '
     SNArray = "" & vbNullChar
     UNArray1 = OldUname & vbNullChar
     UNArray2 = NewUname & vbNullChar
   '
   ' Allocate buffer space
   '
     Result = NetAPIBufferAllocate(UBound(UNArray1) + 1, UNPtr)
     Result = NetAPIBufferAllocate(UBound(UNArray2) + 1, UNPtr)
   
   '
   ' Copy arrays to the buffer
   '
     Result = StrToPtr(UNPtr, UNArray2(0))
   '
   ' Fill the structure
   '
    UserStruct.ptrName = UNPtr
     
     Result = NetUserSetInfo2(SNArray(0), UNArray1(0), 0, UserStruct, ParmError)
     RenameUser = Result
     'MsgBox ("res:" & Result & " , param :" & ParmError)
     Result = NetAPIBufferFree(UNPtr)
    Exit Function
Errorhandler:
     
     RenameUser = CLng(Err.Number)

   End Function

Function DelUser(ByVal Uname As String) As Long
   
   Dim UNArray() As Byte, SNArray() As Byte
     UNArray = Uname & vbNullChar
     SNArray = "" & vbNullChar
     DelUser = NetUserDel(SNArray(0), UNArray(0))
    
    'DelUser = NetUserDel("" & vbNullChar, Uname & vbNullChar)

End Function

