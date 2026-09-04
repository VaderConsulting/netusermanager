Attribute VB_Name = "NTUsers_Bas"
 Option Explicit
   Option Base 0     ' Important assumption for this code
'---------------------------------------------------------------------------------------------------------
'---------------------------DECLARACIONES PARA PERMISOS USER-FOLDER--------------------
'---------------------------------------------------------------------------------------------------------
Declare Function GetFileSecurity Lib "advapi32.dll" Alias "GetFileSecurityA" (ByVal lpFileName As String, ByVal RequestedInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function LookupAccountName Lib "advapi32.dll" Alias "LookupAccountNameA" (ByVal lpSystemName As String, ByVal lpAccountName As String, Sid As Long, cbSid As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Integer) As Long
Declare Function InitializeSecurityDescriptor Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal dwRevision As Long) As Long
Declare Function GetSecurityDescriptorDacl Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, lpbDaclPresent As Long, pDacl As ACL, lpbDaclDefaulted As Long) As Long
Declare Function GetAclInformation Lib "advapi32.dll" (pAcl As ACL, pAclInformation As Any, ByVal nAclInformationLength As Long, ByVal dwAclInformationClass As Integer) As Long
Declare Function LocalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
Declare Function InitializeAcl Lib "advapi32.dll" (pAcl As ACL, ByVal nAclLength As Long, ByVal dwAclRevision As Long) As Long
Declare Function GetAce Lib "advapi32.dll" (pAcl As ACL, ByVal dwAceIndex As Long, pAce As Any) As Long
Declare Function AddAce Lib "advapi32.dll" (pAcl As ACL, ByVal dwAceRevision As Long, ByVal dwStartingAceIndex As Long, pAceList As Any, ByVal nAceListLength As Long) As Long
Declare Function AddAccessAllowedAce Lib "advapi32.dll" (pAcl As ACL, ByVal dwAceRevision As Long, ByVal AccessMask As Long, pSid As Any) As Long
Declare Function SetSecurityDescriptorDacl Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal bDaclPresent As Long, pDacl As ACL, ByVal bDaclDefaulted As Long) As Long
Declare Function SetFileSecurity Lib "advapi32.dll" Alias "SetFileSecurityA" (ByVal lpFileName As String, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long

Declare Function GetLastError Lib "kernel32" () As Long

Public Const SECURITY_DESCRIPTOR_MIN_LENGTH = (20)
Public Const CARACTERES_ERRONEOS = """/\[]:;|=,.+*?<>"
'__________________
Const TABLE = "<TABLE BORDER=""1"" WIDTH=""100%""  CELLSPACING=""1"" CELLPADDING=""3"">" & vbCrLf
Const TR = vbCrLf & "<TR>" & vbCrLf
Const TDa = "<TD BGCOLOR="""
Const TDb = """ WIDTH="""
Const TDc = "%"" ALIGN="""
Const TDd = """ >"
Const TR2 = "</TR>" & vbCrLf
Const TD2 = "</TD>" & vbCrLf

'__________________
Type ACL
        AclRevision As Byte
        Sbz1 As Byte
        AclSize As Integer
        AceCount As Integer
        Sbz2 As Integer
End Type
Type SECURITY_DESCRIPTOR
        Revision As Byte
        Sbz1 As Byte
        Control As Long
        Owner As Long
        Group As Long
        Sacl As ACL
        Dacl As ACL
End Type
Type ACL_SIZE_INFORMATION
        AceCount As Long
        AclBytesInUse As Long
        AclBytesFree As Long
End Type

'---------------------------------------------------------------------------------------------------------
'---------------------------DECLARACIONES PARA CREAR USERS Y GRUPOS--------------------
'---------------------------------------------------------------------------------------------------------
   Type MungeLong
     X As Long
     Dummy As Integer
   End Type

   Type MungeInt
     XLo As Integer
     XHi As Integer
     Dummy As Integer
   End Type

   Type TUser0                    ' Level 0
     ptrName As Long
   End Type

   Type TUser1                    ' Level 1
     ptrName As Long                    '    LPWSTR   usri1_name;
     ptrPassword As Long                'LPWSTR   usri1_password;
     dwPasswordAge As Long          'DWORD  usri1_password_age;
     dwPriv As Long                         'DWORD    usri1_priv;
     ptrHomeDir As Long                 'LPWSTR   usri1_home_dir;
     ptrComment As Long                 'LPWSTR   usri1_comment;
     dwFlags As Long                        'DWORD    usri1_flags;
     ptrScriptPath As Long              'LPWSTR   usri1_script_path;
   End Type
   Type LocalGroup3             '   typedef struct _LOCALGROUP_MEMBERS_INFO_3 {
    dominioynombre As Long  '  LPWSTR       lgrmi3_domainandname;
    End Type
   
   '
   ' for dwPriv
   '
   Global Const USER_PRIV_MASK = &H3
   Global Const USER_PRIV_GUEST = &H0
   Global Const USER_PRIV_USER = &H1
   Global Const USER_PRIV_ADMIN = &H2

   '
   ' for dwFlags
   '
   Global Const UF_SCRIPT = &H1
   Global Const UF_ACCOUNTDISABLE = &H2
   Global Const UF_HOMEDIR_REQUIRED = &H8
   Global Const UF_LOCKOUT = &H10
   Global Const UF_PASSWD_NOTREQD = &H20
   Global Const UF_PASSWD_CANT_CHANGE = &H40
   Global Const UF_NORMAL_ACCOUNT = &H200     ' Needs to be ORed with the
                                       ' other flags
   Global Const UF_DONT_EXPIRE_PASSWD = &H10000

   '
   ' for lFilter
   '
   Global Const FILTER_NORMAL_ACCOUNT = &H2

   Declare Function NetGetDCName Lib "NETAPI32.DLL" (ServerName As Byte, _
   DomainName As Byte, DCNPtr As Long) As Long

   Declare Function NetUserDel Lib "NETAPI32.DLL" (ServerName As Byte, _
   UserName As Byte) As Long
    'Declare Function NetUserDel Lib "NETAPI32.DLL" (ByVal ServerName As String, _
                                                    ByVal UserName As String) _
                                                    As Long
   
   Declare Function NetGroupAddUser Lib "NETAPI32.DLL" (ServerName As _
   Byte, GroupName As Byte, UserName As Byte) As Long
   
  ' NetLocalGroupAddMembers ( _
    IN  LPCWSTR     servername OPTIONAL, _
    IN  LPCWSTR     groupname, _
    IN  DWORD      level, _
    IN  LPBYTE     buf, _
    IN  DWORD      totalentries
    Declare Function NetLocalGroupAddMembers Lib "NETAPI32.DLL" (ServerName As Byte, _
                                                             GroupName As Byte, _
                                                             ByVal Level As Long, _
                                                             Buffer As LocalGroup3, _
                                                             ByVal TotalEntries As Long) _
                                                             As Long

   Declare Function NetGroupDelUser Lib "NETAPI32.DLL" (ServerName As Byte, GroupName As Byte, UserName As Byte) As Long
   
   'Declare Function NetUserChangePassword Lib "NETAPI32.DLL" (ByVal DomainName As String, _
                                                        ByVal UserName As String, _
                                                        ByVal OldPwd As String, _
                                                        ByVal NewPwd As String) As Long
    Declare Function NetUserChangePassword Lib "NETAPI32.DLL" (DomainName As Byte, _
                                                        UserName As Byte, _
                                                        OldPwd As Byte, _
                                                        NewPwd As Byte) As Long

   ' Add using Level 1 user structure
     'NetUserAdd ( _
    IN  LPCWSTR     servername OPTIONAL, _
    IN  DWORD      level, _
    IN  LPBYTE     buf, _
    OUT LPDWORD    parm_err OPTIONAL  );
Declare Function NetUserAdd1 Lib "NETAPI32.DLL" Alias "NetUserAdd" _
   (ServerName As Byte, ByVal Level As Long, Buffer As TUser1, ParmError _
   As Long) As Long
Declare Function NetUserSetInfo Lib "NETAPI32.DLL" _
   (ServerName As Byte, UserName As Byte, ByVal Level As Long, _
   Buffer As TUser1, ParmError As Long) As Long
Declare Function NetUserSetInfo2 Lib "NETAPI32.DLL" Alias "NetUserSetInfo" _
   (ServerName As Byte, UserName As Byte, ByVal Level As Long, _
   Buffer As TUser0, ParmError As Long) As Long

   ' Enumerate using Level 0 user structure
   Declare Function NetUserEnum0 Lib "NETAPI32.DLL" Alias "NetUserEnum" _
   (ServerName As Byte, ByVal Level As Long, ByVal lFilter As Long, _
   Buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, _
   TotalEntries As Long, ResumeHandle As Long) As Long

   Declare Function NetGroupEnumUsers0 Lib "NETAPI32.DLL" Alias _
   "NetGroupGetUsers" (ServerName As Byte, GroupName As Byte, _
   ByVal Level As Long, Buffer As Long, ByVal PrefMaxLen As Long, _
   EntriesRead As Long, TotalEntries As Long, ResumeHandle As Long) As Long

   Declare Function NetGroupEnum0 Lib "NETAPI32.DLL" Alias "NetGroupEnum" _
   (ServerName As Byte, ByVal Level As Long, Buffer As Long, ByVal _
   PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long, _
   ResumeHandle As Long) As Long

   Declare Function NetUserGetGroups0 Lib "NETAPI32.DLL" Alias _
   "NetUserGetGroups" (ServerName As Byte, UserName As Byte, _
   ByVal Level As Long, Buffer As Long, ByVal PrefMaxLen As Long, _
   EntriesRead As Long, TotalEntries As Long) As Long

   Declare Function NetAPIBufferFree Lib "NETAPI32.DLL" Alias _
   "NetApiBufferFree" (ByVal Ptr As Long) As Long

   Declare Function NetAPIBufferAllocate Lib "NETAPI32.DLL" Alias _
   "NetApiBufferAllocate" (ByVal ByteCount As Long, Ptr As Long) As Long

   Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyW" _
   (RetVal As Byte, ByVal Ptr As Long) As Long

   Declare Function StrToPtr Lib "kernel32" Alias "lstrcpyW" _
   (ByVal Ptr As Long, Source As Byte) As Long

   Declare Function PtrToInt Lib "kernel32" Alias "lstrcpynW" _
   (RetVal As Any, ByVal Ptr As Long, ByVal nCharCount As Long) As Long

   Declare Function StrLen Lib "kernel32" Alias "lstrlenW" _
   (ByVal Ptr As Long) As Long


 Function AddAccessRights(ByVal UserName As String, _
                                        ByVal FileName As String, _
                                        ByVal Mascara As Long) As Boolean


'-------------------------------------------------------------------------------------------------------
'------------------------------------COMIENZA LA FIESTA----------------------------------------
'-------------------------------------------------------------------------------------------------------
Dim Res As Long
Dim dwUserNameLength As Long
Dim DomainName As String
Dim SystemName As String
Dim psnuType As Integer
Dim UserSID As Long, cbSid As Long
Dim DomRef As Long
   
   'Res = GetUserName(lpszUserName, dwUserNameLength)
   
   '{
   '   printf("Error %d:GetUserName\n",GetLastError());
   '   return(FALSE);
   '}

   '// STEP 2: Get SID for current user
    SystemName = ""
    DomainName = ""
    UserSID = 0
    psnuType = 0
   Res = LookupAccountName(SystemName, _
                           UserName, _
                           UserSID, _
                           cbSid, _
                           DomainName, _
                           DomRef, _
                           psnuType)
   
   If Res = 0 Then Res = GetLastError()
   
   '// STEP 3: Get security descriptor (SD) for file

   'if(!GetFileSecurity(pFileName,
   '              (SECURITY_INFORMATION)(DACL_SECURITY_INFORMATION),
   '              pFileSD,
   '              SD_SIZE,
   '              (LPDWORD)&dwSDLengthNeeded))
   '{
   '   printf("Error %d:GetFileSecurity\n",GetLastError());
   '   return(FALSE);
   '}

   '// STEP 4: Initialize new SD

   'if(!InitializeSecurityDescriptor(psdNewSD,SECURITY_DESCRIPTOR_REVISION))
   '{
   '   printf("Error %d:InitializeSecurityDescriptor\n",GetLastError());
   '   return(FALSE);
   '}

   '// STEP 5: Get DACL from SD

   'if (!GetSecurityDescriptorDacl(pFileSD,
   '                 &bDaclPresent,
   '                 &pACL,
   '                 &bDaclDefaulted))
   '{
    '  printf("Error %d:GetSecurityDescriptorDacl\n",GetLastError());
   '   return(FALSE);
   '}

   '// STEP 6: Get file ACL size information

   'if(!GetAclInformation(pACL,&AclInfo,sizeof(ACL_SIZE_INFORMATION),
   '   AclSizeInformation))
   '{
   '   printf("Error %d:GetAclInformation\n",GetLastError());
   '   return(FALSE);
   '}

   '// STEP 7: Compute size needed for the new ACL

   'dwNewACLSize = AclInfo.AclBytesInUse +
   '               sizeof(ACCESS_ALLOWED_ACE) +
   '               GetLengthSid(UserSID) - sizeof(DWORD);

   '// STEP 8: Allocate memory for new ACL

   'pNewACL = (PACL)LocalAlloc(LPTR, dwNewACLSize);

   '// STEP 9: Initialize the new ACL

   'if(!InitializeAcl(pNewACL, dwNewACLSize, ACL_REVISION2))
   '{
   '   printf("Error %d:InitializeAcl\n",GetLastError());
   '   LocalFree((HLOCAL) pNewACL);
   '   return(FALSE);
   '}

   '// STEP 10: If DACL is present, copy it to a new DACL

   'if(bDaclPresent)  // only copy if DACL was present
   '{
   '   // STEP 11: Copy the file's ACEs to our new ACL

   '   if(AclInfo.AceCount)
   '   {
   '      for(CurrentAceIndex = 0; CurrentAceIndex < AclInfo.AceCount;
   '         CurrentAceIndex++)
   '      {
   '         // STEP 12: Get an ACE

            'if(!GetAce(pACL,CurrentAceIndex,&pTempAce))
            '{
            '  printf("Error %d: GetAce\n",GetLastError());
            '  LocalFree((HLOCAL) pNewACL);
            '  return(FALSE);
            '}

             '// STEP 13: Add the ACE to the new ACL

            'if(!AddAce(pNewACL, ACL_REVISION, MAXDWORD, pTempAce,
            '   ((PACE_HEADER)pTempAce)->AceSize))
            '{
               'printf("Error %d:AddAce\n",GetLastError());
               'LocalFree((HLOCAL) pNewACL);
               'return(FALSE);
            '}
          '}
      '}
   '}

   '// STEP 14: Add the access-allowed ACE to the new DACL

   'if(!AddAccessAllowedAce(pNewACL,ACL_REVISION2,dwAcessMask, &UserSID))
   '{
   '   printf("Error %d:AddAccessAllowedAce",GetLastError());
   '   LocalFree((HLOCAL) pNewACL);
   '   return(FALSE);
   '}

   '// STEP 15: Set our new DACL to the file SD

   'if (!SetSecurityDescriptorDacl(psdNewSD,
   '                  TRUE,
   '                  pNewACL,
   '                  FALSE))
   '{
   '   printf("Error %d:SetSecurityDescriptorDacl",GetLastError());
   '   LocalFree((HLOCAL) pNewACL);
   '   return(FALSE);
   '}

   '// STEP 16: Set the SD to the File

   'if (!SetFileSecurity(pFileName, DACL_SECURITY_INFORMATION,psdNewSD))
   '{
   '   printf("Error %d:SetFileSecurity\n",GetLastError());
   '   LocalFree((HLOCAL) pNewACL);
   '   return(FALSE);
   '}

   '// STEP 17: Free the memory allocated for the new ACL

   'LocalFree((HLOCAL) pNewACL);
   'return(TRUE);



'}


End Function
    
    

Function ConFont(st As String, _
                    Optional Sz, _
                    Optional Tp, _
                    Optional Cl) As String

If IsMissing(Sz) Then Sz = "2"
If IsMissing(Tp) Then Tp = "Arial Black"
If IsMissing(Cl) Then Cl = "Navy"

ConFont = "<FONT SIZE=""" & Sz & _
            """ FACE=""" & Tp & _
            """ COLOR=""" & Cl & _
            """>" & st & "</FONT>"

End Function


   Function DelUserFromGroup(ByVal SName As String, _
   ByVal GName As String, ByVal Uname As String) As Long
   '
   ' This only deletes users from global groups - not local groups
   '
   Dim SNArray() As Byte, GNArray() As Byte, UNArray() As Byte, _
   Result As Long
     SNArray = SName & vbNullChar
     GNArray = GName & vbNullChar
     UNArray = Uname & vbNullChar
     Result = NetGroupDelUser(SNArray(0), GNArray(0), UNArray(0))
     If Result = 2220 Then Debug.Print _
   "There is no **GLOBAL** group '" & GName & "'"
     DelUserFromGroup = Result
   End Function

Function EnCelda(st As String, _
                Numcelda As Integer, _
                Optional ElColor, _
                Optional ElWidth, _
                Optional ElAlign) As String

If IsMissing(ElColor) Then ElColor = "#FCFCED"
If IsMissing(ElWidth) Then ElWidth = "*"
If IsMissing(ElAlign) Then ElAlign = "Left"

ElColor = CStr(ElColor)
ElWidth = CStr(ElWidth)
ElAlign = CStr(ElAlign)

Select Case Numcelda
    Case 1 'Celda primera
        EnCelda = TR & TDa & ElColor & TDb & ElWidth & TDc & ElAlign & TDd & st & TD2
    Case 2 'Ni primera ni ultima
        EnCelda = TDa & ElColor & TDb & ElWidth & TDc & ElAlign & TDd & st & TD2
    Case 3 'Celda ultima
        EnCelda = TDa & ElColor & TDb & ElWidth & TDc & ElAlign & TDd & st & TD2 & TR2
    Case 4 'Celda ultima y primera a la vez
        EnCelda = TR & TDa & ElColor & TDb & ElWidth & TDc & ElAlign & TDd & st & TD2 & TR2
End Select

End Function
Function EnTabla(st As String) As String

EnTabla = TABLE & st & "</table>"

End Function


   Function EnumerateGroups(ByVal SName As String, _
   ByVal Uname As String) As Long
   '
   ' Enumerates global groups only - not local groups
   '
   ' The buffer is filled from the left with pointers to user names that
   ' are filled from the right side. For example:
   '
   '     ptr1|ptr2|...|ptrn|<garbage>|strn|...|str2|str1
   '     ^-------------- BufPtr buffer ----------------^
   '
   ' On NT, TotalEntries is the number of entries left to be read including
   ' the currently read entries.
   '
   ' On LanMan and OS/2, it is the total number of entries, period. Code
   ' would have to be changed to reflect this if the Domain controller
   ' wasn't an NT machine.
   '
   ' BufPtr gets the address of the buffer (or ptr1 - add 4 to BufPtr for
   ' each additional pointer)
   '
   Dim Result As Long, BufPtr As Long, EntriesRead As Long, _
   TotalEntries As Long, ResumeHandle As Long, BufLen As Long, _
   SNArray() As Byte, GNArray(99) As Byte, UNArray() As Byte, _
   GName As String, I As Integer, UNPtr As Long, _
   TempPtr As MungeLong, TempStr As MungeInt

     SNArray = SName & vbNullChar       ' Move to byte array
     UNArray = Uname & vbNullChar       ' Move to Byte array
     BufLen = 255                       ' Buffer size
     ResumeHandle = 0                   ' Start with the first entry

     Do
       If Uname = "" Then
         Result = NetGroupEnum0(SNArray(0), 0, BufPtr, BufLen, _
   EntriesRead, TotalEntries, ResumeHandle)
       Else
         Result = NetUserGetGroups0(SNArray(0), UNArray(0), 0, BufPtr, _
   BufLen, EntriesRead, TotalEntries)
       End If
       EnumerateGroups = Result
       If Result <> 0 And Result <> 234 Then    ' 234 means multiple reads
                                                ' required
         Debug.Print "Error " & Result & " enumerating group " & _
   EntriesRead & " of " & TotalEntries
         Exit Function
       End If
       For I = 1 To EntriesRead
         ' Get pointer to string from beginning of buffer
         ' Copy 4 byte block of memory in 2 steps
         Result = PtrToInt(TempStr.XLo, BufPtr + (I - 1) * 4, 2)
         Result = PtrToInt(TempStr.XHi, BufPtr + (I - 1) * 4 + 2, 2)
         LSet TempPtr = TempStr ' munge 2 Integers to a Long
         ' Copy string to array and convert to a string
         Result = PtrToStr(GNArray(0), TempPtr.X)
         GName = Left(GNArray, StrLen(TempPtr.X))
         Debug.Print "Group: " & GName
       Next I
     Loop Until EntriesRead = TotalEntries
   ' The above condition only valid for reading accounts on NT
   ' but not OK for OS/2 or LanMan

     Result = NetAPIBufferFree(BufPtr)         ' Don't leak memory

   End Function

   Function EnumerateUsers(ByVal SName As String, ByVal GName As String) _
   As Long
   '
   ' If a group name is specified, it must be a global group
   ' and not a local group.
   '
   ' The buffer is filled from the left with pointers to user names that
   ' are filled from the right side. For example:
   '
   '     ptr1|ptr2|...|ptrn|<garbage>|strn|...|str2|str1
   '     ^-------------- BufPtr buffer ----------------^
   '
   ' On Windows NT, TotalEntries is the number of entries left to be read,
   ' including the currently read entries.
   ' On LanMan and OS/2, it is the total number of entries, period.  Code
   ' would have to be changed to reflect this if the Domain controller
   ' wasn't an NT machine.
   '
   ' BufPtr gets the address of the buffer (or ptr1 - add 4 to BufPtr for
   ' each additional pointer)
   '
   ' SName should be "\\servername"
   '
   Dim Result As Long, BufPtr As Long, EntriesRead As Long, _
   TotalEntries As Long, ResumeHandle As Long, BufLen As Long, _
   SNArray() As Byte, GNArray() As Byte, UNArray(99) As Byte, _
   Uname As String, I As Integer, UNPtr As Long, TempPtr As MungeLong, _
   TempStr As MungeInt

     SNArray = SName & vbNullChar       ' Move to byte array
     GNArray = GName & vbNullChar       ' Move to Byte array
     BufLen = 255                       ' Buffer size
     ResumeHandle = 0                   ' Start with the first entry

     Do
       If GName = "" Then
         Result = NetUserEnum0(SNArray(0), 0, FILTER_NORMAL_ACCOUNT, _
   BufPtr, BufLen, EntriesRead, TotalEntries, ResumeHandle)
       Else
         Result = NetGroupEnumUsers0(SNArray(0), GNArray(0), 0, BufPtr, _
   BufLen, EntriesRead, TotalEntries, ResumeHandle)
       End If
       EnumerateUsers = Result
       If Result <> 0 And Result <> 234 Then    ' 234 means multiple reads
                                                ' required
         Debug.Print "Error " & Result & " enumerating user " _
   & EntriesRead & " of " & TotalEntries
         If Result = 2220 Then Debug.Print _
   "There is no **GLOBAL** group '" & GName & "'"
         Exit Function
       End If
       For I = 1 To EntriesRead
         ' Get pointer to string from beginning of buffer
         ' Copy 4-byte block of memory in 2 steps
         Result = PtrToInt(TempStr.XLo, BufPtr + (I - 1) * 4, 2)
         Result = PtrToInt(TempStr.XHi, BufPtr + (I - 1) * 4 + 2, 2)
         LSet TempPtr = TempStr ' munge 2 integers into a Long
         ' Copy string to array
         Result = PtrToStr(UNArray(0), TempPtr.X)
         Uname = Left(UNArray, StrLen(TempPtr.X))
         Debug.Print "User: " & Uname
       Next I
     Loop Until EntriesRead = TotalEntries
   ' The above condition is only valid for reading accounts on Windows NT,
   ' but is not OK for OS/2 or LanMan

     Result = NetAPIBufferFree(BufPtr)         ' Don't leak memory

   End Function

   Function GetPrimaryDCName(ByVal MName As String, _
   ByVal DName As String) As String
   Dim Result As Long, DCName As String, DCNPtr As Long
   Dim DNArray() As Byte, MNArray() As Byte, DCNArray(100) As Byte
     MNArray = MName & vbNullChar
     DNArray = DName & vbNullChar
     Result = NetGetDCName(MNArray(0), DNArray(0), DCNPtr)
     If Result <> 0 Then
       Debug.Print "Error: " & Result
       Exit Function
     End If
     Result = PtrToStr(DCNArray(0), DCNPtr)
     Result = NetAPIBufferFree(DCNPtr)
     DCName = DCNArray()
     GetPrimaryDCName = DCName
   End Function

Sub Main()
App.HelpFile = "C:\Projects\ApiUsers\FinalApiDef\TSAI_NTUSERS.HLP"
End Sub
Function EsOkUserName(ByVal Uname As String) As Boolean
Dim lon As Integer, t As Integer
Dim it As String

EsOkUserName = True

lon = Len(Uname)
If lon > 0 Then
    For t = 1 To lon
        it = Mid(Uname, t, 1)
        If InStr(CARACTERES_ERRONEOS, it) > 0 Then
            EsOkUserName = False
            Exit For
        End If
    Next
Else
    EsOkUserName = False
End If

End Function

Function NKS(st As String, que As Integer) As String
'1 equivale a Negrita
'2 equivale Cursiva
'4 equivale a Subrayado
Dim Temp$

Temp$ = st
If que - 4 >= 0 Then
    que = que - 4
    Temp$ = "<U>" & st & "</U>"
End If
If que - 2 >= 0 Then
    que = que - 2
    Temp$ = "<K>" & Temp$ & "</K>"
End If
If que - 1 >= 0 Then
    que = que - 1
    Temp$ = "<B>" & Temp$ & "</B>"
End If

    NKS = Temp$
    
End Function


