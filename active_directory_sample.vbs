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
WScript.Echo "姓: " & objUser.sn
WScript.Echo "名: " & objUser.givenName
'WScript.Echo "First name: " & objUser.FirstName
WScript.Echo "イニシャル: " & objUser.initials
 
'WScript.Echo "Last name: " & objUser.LastName
WScript.Echo "表示名: " & objUser.displayName
'WScript.Echo "Display name: " & objUser.FullName
WScript.Echo "説明: " & objUser.description
WScript.Echo "事業所: " & objUser.physicalDeliveryOfficeName
WScript.Echo "電話番号: " & objUser.telephoneNumber
WScript.Echo "その他の電話番号: " & objUser.otherTelephone
WScript.Echo "E-メール: " & objUser.mail
' WScript.Echo "Email: " & objUser.EmailAddress
WScript.Echo "Webページ: " & objUser.wWWHomePage
WScript.Echo "その他のWebページ: " & objUser.url
WScript.Echo ""
WScript.Echo "住所"
WScript.Echo "番地: " & objUser.streetAddress
WScript.Echo "私書箱: " & objUser.postOfficeBox
WScript.Echo "市区町村: " & objUser.l
WScript.Echo "都道府県: " & objUser.st
WScript.Echo "郵便番号: " & objUser.postalCode
WScript.Echo "国コード: " & objUser.countryCode
'WScript.Echo "Country/region: " & objUser.c    '(ISO 4217)
WScript.Echo ""
WScript.Echo "ACCOUNT"
WScript.Echo "ユーザーログオン名: " & objUser.userPrincipalName
WScript.Echo "ユーザーログオン名(Windows 2000以前): " & objUser.sAMAccountName
WScript.Echo "アカウントは無効: " & objUser.AccountDisabled
' WScript.Echo "Account Control #: " & objUser.userAccountControl
WScript.Echo "ログオン時間: " & objUser.logonHours
WScript.Echo "ログオン先: " & objUser.userWorkstations
' WScript.Echo "User must change password at next logon: " & objUser.pwdLastSet
WScript.Echo "ユーザーはパスワードを変更できない: " & objUser.userAccountControl
WScript.Echo "パスワードを無期限にする: " & objUser.userAccountControl
WScript.Echo "暗号化を元に戻せる状態でパスワードを保存する: " & objUser.userAccountControl
' WScript.Echo "Account expires end of (date): " & objUser.accountExpires
WScript.Echo ""
WScript.Echo "プロファイル"
WScript.Echo "プロファイル パス: " & objUser.profilePath
' WScript.Echo "Profile path: " & objUser.Profile
WScript.Echo "ログオンスクリプト: " & objUser.scriptPath
WScript.Echo "Home folder, local path: " & objUser.homeDirectory
WScript.Echo "Home folder, Connect, Drive: " & objUser.homeDrive
WScript.Echo "Home folder, Connect, To:: " & objUser.homeDirectory
WScript.Echo ""
WScript.Echo "電話"
WScript.Echo "自宅: " & objUser.homePhone
WScript.Echo "Other Home phone numbers: " & objUser.otherHomePhone
WScript.Echo "ポケットベル: " & objUser.pager
WScript.Echo "Other Pager numbers: " & objUser.otherPager
WScript.Echo "携帯電話: " & objUser.mobile
WScript.Echo "Other Mobile numbers: " & objUser.otherMobile
WScript.Echo "FAX: " & objUser.facsimileTelephoneNumber
WScript.Echo "Other Fax numbers: " & objUser.otherFacsimileTelephoneNumber
WScript.Echo "IP電話: " & objUser.ipPhone
WScript.Echo "Other IP phone numbers: " & objUser.otherIpPhone
WScript.Echo "メモ: " & objUser.info
WScript.Echo ""
WScript.Echo "組織"
WScript.Echo "役職: " & objUser.title
WScript.Echo "部署: " & objUser.department
WScript.Echo "会社: " & objUser.company
WScript.Echo "上司: " & objUser.manager
