<div align="center">

## Get NT User Info \(FullName, Groups\) using ADSI


</div>

### Description

Example code showing how one may extract NT Domain User information using ADSI (Active Directory Service Interfaces).

This code simply extracts a user's FullName and lists the Groups to which he/she belongs, given his username.

This code will work across domains, provided the correct authentication values (username, password) are inserted.
 
### More Info
 
Uses ADSI (Active Directory Services) 2.0 or later.

Example from IMMEDIATE WINDOW:

----

NT UserName    Webdood

FullName      Shannon Norrell

This user belongs to the following NT Groups:

Domain Users


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Shannon Norrell](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/shannon-norrell.md)
**Level**          |Unknown
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/shannon-norrell-get-nt-user-info-fullname-groups-using-adsi__1-3448/archive/master.zip)

### API Declarations

Must Reference the ActiveDS Type Library. (activeds.tlb)


### Source Code

```
Private oIADS As ActiveDs.IADsContainer
Private oUser As ActiveDs.IADsUser
Private oGroup As ActiveDs.IADsGroup
Private Sub Form_Load()
  txtDomain = "MYDOMAIN"
  usrName = "Administrator"
  usrPassword = "sa"
  usrNameOfInterest = "WebDood"
  Set oIADS = GetObject("WinNT:").OpenDSObject("WinNT://" & txtDomain, usrName, usrPassword, 1)
  Set oUser = oIADS.GetObject("user", usrNameOfInterest)
  With oUser
   Debug.Print "NT UserName" & Space$(8) & .Name
   Debug.Print "FullName" & Space$(11) & .FullName
   Debug.Print "This user belongs to the following NT Groups:"
   For Each oGroup In .Groups
     Debug.Print vbTab & oGroup.Name
   Next
  End With
End Sub
```

