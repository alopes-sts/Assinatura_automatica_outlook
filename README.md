# Assinatura automatica outlook
Exemplo de código em vbscript para automatizar a assinatura do outlook + chave no registro para bloqueio, evitando que o usuário possa alterar.

O primeiro arquivo .VBS é um exemplo que foi feito atendendo a solicitação do cliente de que a assinatura deveria ocupar somente uma linha:  Nome | Departamento | Telefone | Celular | SiteDaEmpresa

O segundo arquivo .VBS é uma assinatura mais convêncional, cada campo em uma linha:
Nome 
Departamento
Telefone
Celular
SiteDaEmpresa


_____________________________________________________________________________
No registro do Windows deve-se incluir as 2 chaves abaixo:

Local: HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\MailSettings    
Chave: Assinatura_Empresa

Tipo: REG_SSZ

Chave: ReplySignature

Tipo: REG_SSZ


Notar que 16.0 é a versão do Office, no meu caso, a empresa possuia offices de 3 versões diferentes, então criei as chaves para as 3 versões do office na GPO.
Com essa alteração no registro do Windows o usuário fica impedido de fazer qualquer alteração na assinatura, seja pelo painel de configurações do Outlook ou pelo botão 'INSERIR' dentro da mensagem que está sendo criada.
Atentar para o nome da chave (Assinatura_Empresa), tem que ser o mesmo nome da assinatura
