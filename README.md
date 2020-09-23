<div align="center">

## Create/Destroy User on Domain \( Administrating NT\)


</div>

### Description

Create a new user and destroy an existing user on a Windows NT domain..

When a user is

created, I set him to be a member of Domain Users while you can specify that he goes

into Domain User, Domain Guests, Domain Admins

Hong YAN <HONG-YAN@worldnet.att.net>
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Newsgroup Posting](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/newsgroup-posting.md)
**Level**          |Unknown
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/newsgroup-posting-create-destroy-user-on-domain-administrating-nt__1-532/archive/master.zip)

### API Declarations

```
'Jeff Hong YAN 11/20/96 modified on 4/18/97
'This module shows how to create / destroy a user account.
'Modified according to MS KB article Q159498
'You must have account operator's right to run
' for dwPriv
Const USER_PRIV_MASK = &H3
Const USER_PRIV_GUEST = &H0
Const USER_PRIV_USER = &H1
Const USER_PRIV_ADMIN = &H2
' for dwFlags
Const UF_SCRIPT = &H1
Const UF_ACCOUNTDISABLE = &H2
Const UF_HOMEDIR_REQUIRED = &H8
Const UF_LOCKOUT = &H10
Const UF_PASSWD_NOTREQD = &H20
Const UF_PASSWD_CANT_CHANGE = &H40
Const UF_NORMAL_ACCOUNT = &H200
Declare Function StrToPtr Lib "kernel32" Alias "lstrcpyW" ( _
ByVal Ptr As Long, Source As Byte) As Long
' Add using Level 1 user structure
Declare Function NetUserAdd1 Lib "NETAPI32.DLL" Alias "NetUserAdd" _
(ServerName As Byte, ByVal Level As Long, Buffer As TUser1, lParmError _
As Long) As Long
Declare Function NetUserDel Lib "NETAPI32.DLL" (ServerName As Byte, _
UserName As Byte) As Long
Type TUser1          ' Level 1
 ptrName As Long
 ptrPassword As Long
 dwPasswordAge As Long
 dwPriv As Long
 ptrHomeDir As Long
 ptrComment As Long
 dwFlags As Long
 ptrScriptHomeDir As Long
End Type
Declare Function NetAPIBufferFree Lib "NETAPI32.DLL" Alias _
"NetApiBufferFree" (ByVal Ptr As Long) As Long
Declare Function NetAPIBufferAllocate Lib "NETAPI32.DLL" Alias _
"NetApiBufferAllocate" (ByVal ByteCount As Long, Ptr As Long) As Long
```


### Source Code

```
Function DomainCreateUser( _
  ByVal sSName As String, _
  ByVal sUName As String, _
  ByVal sPWD As String, _
  ByVal sHomeDir As String, _
  ByVal sComment As String, _
  ByVal sScriptFile As String) As Long
'Create a new user to be a member of group Domain Users
  Dim lResult As Long
  Dim lParmError As Long
  Dim lUNPtr As Long
  Dim lPWDPtr As Long
  Dim lHomeDirPtr As Long
  Dim lCommentPtr As Long
  Dim lScriptFilePtr As Long
  Dim bSNArray() As Byte
  Dim bUNArray() As Byte
  Dim bPWDArray() As Byte
  Dim bHomeDirArray() As Byte
  Dim bCommentArray() As Byte
  Dim bScriptFileArray() As Byte
  Dim UserStruct As TUser1
  ' Move to byte arrays
  bSNArray = sSName & vbNullChar
  bUNArray = sUName & vbNullChar
  bPWDArray = sPWD & vbNullChar
  bHomeDirArray = sHomeDir & vbNullChar
  bCommentArray = sComment & vbNullChar
  bScriptFileArray = sScriptFile & vbNullChar
  ' Allocate buffer space
  lResult = NetAPIBufferAllocate(UBound(bUNArray) + 1, lUNPtr)
  lResult = NetAPIBufferAllocate(UBound(bPWDArray) + 1, lPWDPtr)
  lResult = NetAPIBufferAllocate(UBound(bHomeDirArray) + 1, lHomeDirPtr)
  lResult = NetAPIBufferAllocate(UBound(bCommentArray) + 1, lCommentPtr)
  lResult = NetAPIBufferAllocate(UBound(bScriptFileArray) + 1, lScriptFilePtr)
  ' Copy arrays to the buffer
  lResult = StrToPtr(lUNPtr, bUNArray(0))
  lResult = StrToPtr(lPWDPtr, bPWDArray(0))
  lResult = StrToPtr(lHomeDirPtr, bHomeDirArray(0))
  lResult = StrToPtr(lCommentPtr, bCommentArray(0))
  lResult = StrToPtr(lScriptFilePtr, bScriptFileArray(0))
  With UserStruct
   .ptrName = lUNPtr
   .ptrPassword = lPWDPtr
   .dwPasswordAge = 3
   .dwPriv = USER_PRIV_USER
   .ptrHomeDir = lHomeDirPtr
   .ptrComment = lCommentPtr
   .dwFlags = UF_NORMAL_ACCOUNT Or UF_SCRIPT
   .ptrScriptHomeDir = lScriptFilePtr
  End With
  ' Create the new user
  lResult = NetUserAdd1(bSNArray(0), 1, UserStruct, lParmError)
  DomainCreateUser = lResult
  If lResult <> 0 Then
    Call NetErrorHandler(lResult, " when creating new user " & sUName)
  End If
  ' Release buffers from memory
  lResult = NetAPIBufferFree(lUNPtr)
  lResult = NetAPIBufferFree(lPWDPtr)
  lResult = NetAPIBufferFree(lHomeDirPtr)
  lResult = NetAPIBufferFree(lCommentPtr)
  lResult = NetAPIBufferFree(lScriptFilePtr)
End Function
Public Function DomainDestroyUser(ByVal sSName As String, ByVal sUName As String)
'Destroy an existing user with user id sUName
'from current PDC with sSName
  Dim lResult As Long
  Dim lParmError As Long
  Dim bSNArray() As Byte
  Dim bUNArray() As Byte
  ' Move to byte arrays
  bSNArray = sSName & vbNullChar
  bUNArray = sUName & vbNullChar
  lResult = NetUserDel(bSNArray(0), bUNArray(0))
  If lResult = 0 Then
    DomainDestroyUser = True
  Else
    Call NetErrorHandler(lResult, "delete user '" & sUName & "' from server '" &
sSName & "'.")
    DomainDestroyUser = False
  End If
End Function
```

