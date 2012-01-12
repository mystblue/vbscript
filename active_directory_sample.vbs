'userinfo.vbs
 
' Usage:
'       cscript //Nologo userinfo.vbs
 
' List User properties as displayed in ADUC
 
On Error Resume Next
Dim objSysInfo, objUser
Set objSysInfo = CreateObject("ADSystemInfo")
 
' Currently logged in User
WScript.Echo "UserName field:" & objSysInfo.UserName
Set objUser = GetObject("LDAP://" & objSysInfo.UserName)
 ' or specific user:
 'Set objUser = GetObject("LDAP://CN=johndoe,OU=Users,DC=ss64,DC=com")
 
WScript.Echo "DN: " & objUser.distinguishedName
 
WScript.Echo ""
WScript.Echo "GENERAL"
WScript.Echo "��: " & objUser.sn
WScript.Echo "��: " & objUser.givenName
'WScript.Echo "First name: " & objUser.FirstName
WScript.Echo "�C�j�V����: " & objUser.initials
 
'WScript.Echo "Last name: " & objUser.LastName
WScript.Echo "�\����: " & objUser.displayName
'WScript.Echo "Display name: " & objUser.FullName
WScript.Echo "����: " & objUser.description
WScript.Echo "���Ə�: " & objUser.physicalDeliveryOfficeName
WScript.Echo "�d�b�ԍ�: " & objUser.telephoneNumber
WScript.Echo "���̑��̓d�b�ԍ�: " & objUser.otherTelephone
WScript.Echo "E-���[��: " & objUser.mail
' WScript.Echo "Email: " & objUser.EmailAddress
WScript.Echo "Web�y�[�W: " & objUser.wWWHomePage
WScript.Echo "���̑���Web�y�[�W: " & objUser.url
WScript.Echo ""
WScript.Echo "�Z��"
WScript.Echo "�Ԓn: " & objUser.streetAddress
WScript.Echo "������: " & objUser.postOfficeBox
WScript.Echo "�s�撬��: " & objUser.l
WScript.Echo "�s���{��: " & objUser.st
WScript.Echo "�X�֔ԍ�: " & objUser.postalCode
WScript.Echo "���R�[�h: " & objUser.countryCode
'WScript.Echo "Country/region: " & objUser.c    '(ISO 4217)
WScript.Echo ""
WScript.Echo "ACCOUNT"
WScript.Echo "���[�U�[���O�I����: " & objUser.userPrincipalName
WScript.Echo "���[�U�[���O�I����(Windows 2000�ȑO): " & objUser.sAMAccountName
WScript.Echo "�A�J�E���g�͖���: " & objUser.AccountDisabled
' WScript.Echo "Account Control #: " & objUser.userAccountControl
WScript.Echo "���O�I������: " & objUser.logonHours
WScript.Echo "���O�I����: " & objUser.userWorkstations
' WScript.Echo "User must change password at next logon: " & objUser.pwdLastSet
WScript.Echo "���[�U�[�̓p�X���[�h��ύX�ł��Ȃ�: " & objUser.userAccountControl
WScript.Echo "�p�X���[�h�𖳊����ɂ���: " & objUser.userAccountControl
WScript.Echo "�Í��������ɖ߂����ԂŃp�X���[�h��ۑ�����: " & objUser.userAccountControl
' WScript.Echo "Account expires end of (date): " & objUser.accountExpires
WScript.Echo ""
WScript.Echo "�v���t�@�C��"
WScript.Echo "�v���t�@�C�� �p�X: " & objUser.profilePath
' WScript.Echo "Profile path: " & objUser.Profile
WScript.Echo "���O�I���X�N���v�g: " & objUser.scriptPath
WScript.Echo "Home folder, local path: " & objUser.homeDirectory
WScript.Echo "Home folder, Connect, Drive: " & objUser.homeDrive
WScript.Echo "Home folder, Connect, To:: " & objUser.homeDirectory
WScript.Echo ""
WScript.Echo "�d�b"
WScript.Echo "����: " & objUser.homePhone
WScript.Echo "Other Home phone numbers: " & objUser.otherHomePhone
WScript.Echo "�|�P�b�g�x��: " & objUser.pager
WScript.Echo "Other Pager numbers: " & objUser.otherPager
WScript.Echo "�g�ѓd�b: " & objUser.mobile
WScript.Echo "Other Mobile numbers: " & objUser.otherMobile
WScript.Echo "FAX: " & objUser.facsimileTelephoneNumber
WScript.Echo "Other Fax numbers: " & objUser.otherFacsimileTelephoneNumber
WScript.Echo "IP�d�b: " & objUser.ipPhone
WScript.Echo "Other IP phone numbers: " & objUser.otherIpPhone
WScript.Echo "����: " & objUser.info
WScript.Echo ""
WScript.Echo "�g�D"
WScript.Echo "��E: " & objUser.title
WScript.Echo "����: " & objUser.department
WScript.Echo "���: " & objUser.company
WScript.Echo "��i: " & objUser.manager
