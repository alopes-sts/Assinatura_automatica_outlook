On Error Resume Next

'Exemplo de assinatura em uma linha só, a pedido da empresa

set FSO = CreateObject("Scripting.FileSystemObject")
Set wshShell = CreateObject("WScript.Shell")

'Remove assinaturas existentes no perfil do usuário
strUserProfile = wshShell.ExpandEnvironmentStrings( "%USERPROFILE%" )
FSO.DeleteFolder(strUserProfile & "\AppData\Roaming\Microsoft\Signatures")
FSO.DeleteFolder(strUserProfile & "\AppData\Roaming\Microsoft\Assinaturas")

Set objSysInfo = CreateObject("ADSystemInfo")
Set WshShell = CreateObject("WScript.Shell")

strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)

'Define quais os campos serão usados na assinatura (Nome completo, Departamento, Ramal, Celular e Site da empresa)
strName = objUser.FullName
strDepartment = objUser.Department
strPhone = objUser.TelephoneNumber
strMobile = objUser.Mobile
strWeb = objuser.wWWHomePage

Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

'Definições de fonte, cor, tamanho
objSelection.Font.Name = "Verdana"
objSelection.Font.Size = 8
objSelection.Font.Bold = True
objSelection.Font.Color = RGB(0,0,150)

'Criando assinatura
objSelection.TypeText strName
objSelection.Font.Bold = False
objSelection.TypeText " | " & strDepartment & " | Tel.: " & strPhone
'foi colocado um "if" no campo do celular, pois nem todos os colaboradores possuem
if len(strMobile) > 0 then
	objSelection.TypeText " | Cel.: " & strMobile
end if

'Criando um hiperlink para o site da empresa
objSelection.TypeText " | "
objSelection.Hyperlinks.Add objSelection.Range, "" & strWeb & "", , , "www.sitedaempresa.com.br"

Set objSelection = objDoc.Range()

'Definindo o nome da assinatura e removendo o seu uso em respostas de emails
objSignatureEntries.Add "Assinatura_Empresa", objSelection
objSignatureObject.NewMessageSignature = "Assinatura_Empresa"
objSignatureObject.ReplyMessageSignature = "none"

objDoc.Saved = True
objWord.Quit