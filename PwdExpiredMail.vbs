'On Error Resume Next

Const ADS_UF_DONT_EXPIRE_PASSWD = &h10000
Const E_ADS_PROPERTY_NOT_FOUND  = &h8000500D
Const ONE_HUNDRED_NANOSECOND    = .000000100
Const SECONDS_IN_DAY            = 86400

Const useLocalSMTPService = 1
Const useRemoteSMTPService = 2
Const useLocalExchangeService = 3

Const SMTP_SERVER_ADDRESS = "172.17.0.202"
Const SMTP_SERVER_PORT    = 25

Const NOTIFY_DAYS_AGO     = 15
Const ADMIN_EMAIL_ADDRESS = "miura@finos.hakata.fukuoka.jp"
'------------------------------------------------------------------------------
' 
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' �p�X���[�h�|���V�[�擾�imaxPwdAge)
'------------------------------------------------------------------------------
Function GetMaxPwdAge()
    Set objDomain = GetObject("LDAP://DC=FINOS")
    Set GetMaxPwdAge = objDomain.Get("maxPwdAge")
End Function

'------------------------------------------------------------------------------
' nano�b��day�ϊ�
'------------------------------------------------------------------------------
Function getDays(objNano)
    dblNano = Abs(objNano.HighPart * 2^32 + objNano.LowPart)
    dblSecs = dblNano * ONE_HUNDRED_NANOSECOND
    getDays = Int(dblSecs / SECONDS_IN_DAY)
End Function


'------------------------------------------------------------------------------
' Main����
'------------------------------------------------------------------------------
Sub MainProc()

    '�p�X���[�h�L�������擾
    Set objMaxPwdAge = GetMaxPwdAge
    If objMaxPwdAge.LowPart = 0 Then
        WScript.Echo "�p�X���[�h�̍Œ��L�����Ԃ̓h���C������ 0 �ɐݒ肳��Ă��܂��B" & _
                     "���������āA�p�X���[�h�ɗL�������͂���܂���B"
        Exit Sub
    End If        

    'P�����o�[���擾
    Set objConnection = CreateObject("ADODB.Connection")
    objConnection.Open "Provider=ADsDSOObject;"
    Set objCommand = CreateObject("ADODB.Command")
    objCommand.ActiveConnection = objConnection
    objCommand.CommandText = "<LDAP://dc=FINOS>;(&(objectCategory=User)(givenname=*��));sn,givenname,samaccountname,displayName,distinguishedName;subtree"
    Set objRecordset = objCommand.Execute

    Dim strAdmMailBody
    Dim sendCnt
    strAdmMailBody = ""
    sendCnt = 0




    objRecordset.MoveFirst
    While Not objRecordset.EOF
        '���[�U�[���擾
        strUserDN = objRecordset.fields("distinguishedName")
        Set objUser = GetObject("LDAP://" & strUserDN)  

	WScript.echo "sn:" & objUser.sn
	WScript.echo "givenname:" & objUser.givenname
	WScript.echo "samaccountname:" & objUser.samaccountname
	WScript.echo "displayName:" & objUser.displayName
	WScript.echo "distinguishedName:" & objUser.distinguishedName
	WScript.echo "distinguishedName:" & objUser.distinguishedName

        objRecordset.MoveNext
    Wend

    objConnection.Close
    Set objConnection = Nothing
    Set objCommand = Nothing
    Set objRecordset = Nothing
End Sub


MainProc
