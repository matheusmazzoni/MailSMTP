# MailSMTP

Projeto de uma Dll criada em C# para fazer envios de email urilizando um servidor SMTP.

----------

## Primeiros passos:

### Adicionando dll ao sistema (vba):

1. Pegue a MailSMTP.dll e
2. Copie para sua aplicação a pasta src, na qual contem todos as classes que serão utilizadas;
3. Abra o seu projeto e importe a pasta copiada.
4.  A aplicação utiliza o package 'Newtonsoft.Json' para facilitar a deserialização do JSON, respectivamente. 

**Pronto!** Agora, você já pode consumir a NS DDF-e API através do seu sistema. Todas as funcionalidades de comunicação foram implementadas na classe DDFeAPI.cs. Caso tenha dúvidas de como adicionar este package, veja o tutorial a seguir: [Package Newtonsoft.Json](https://docs.microsoft.com/pt-br/nuget/consume-packages/install-use-packages-visual-studio#finding-and-installing-a-package).

----------

## Utilizando o MailSMTP:

### Realizando um envio de email:

Para realizar um envio de email, você poderá utilizar a função **sendEmail** da dll. Veja abaixo sobre os parâmetros necessários, e um exemplo de chamada do método.

#### Parâmetros:

**ATENÇÃO:** as '{}' representadas na tabela de parametros significam de maneira visual um ArrayList.

Parametros     	   | Tipo de Dado | Descrição
:-----------------:|:------------:|:-----------
mailServer         | String       | Endereço do servidor de email smtp; Ex.: "mail.mazzoni.com.br".
mailPort           | Int          | Porta do servidor de email.
fromName           | String   	  | Nome de do remetente.
fromEmailAddress   | String		  | Endereço de email do remetente.
recipients         | ArrayList    | Email(s) e nome(s) do(s) destinatario(s). Deve ser preenchido separados por ';'. Ex.:{"exemplo@gmail.com;Nome Exemplo", "exemplo1@gmail.com;Nome Exemplo1"}
emailSubject       | String       | Assunto do email que será enviado.
emailBody          | String       | Conteudo escrito do email. 
reqAuthentication  | bool         | (Optional)Se é necessario autenticação para utilizar o servidor de email.
userNameSSL        | String       | (Optional) Nome do usuario a se autenticar.
passwordSSL        | String       | (Optional) Senha do usuario a se autenticar.
attachments        | ArrayList    | (Optional) Caminhos dos anexos que deseja enviar. Ex.:{"C:\exemplo.pdf", "C:\exemplo.xml"}.

#### Exemplo de chamada:

Após ter todos os parâmetros listados acima, você deverá fazer a chamada da função. Veja o código em VB6 de exemplo abaixo:
    
	'Cria uma refencia a dll      
    Dim mailSmtp As New MailSMTP.email
	Dim feitoEnvio As Boolean
	
	'Cria o ArrayList de Destinatarios
	Dim dests As Object
    Set dests = CreateObject("System.Collections.ArrayList")
    dests.Add "exemplo@gmail.com;Nome Exemplo"
    dests.Add "exemplo1@gmail.com;Nome Exemplo1"
    
	'Cria o ArrayList de anexos
    Dim anexos As Object
    Set anexos = CreateObject("System.Collections.ArrayList")
    anexos.Add "C:\exemplo.xml"
	anexos.Add "C:\exemplo.pdf"
	
	'Faz o envio de email
    feitoEnvio = mailSmtp.sendEmail("mail.mazzoni.com.br", 26, "Empresa teste", "exemplo@mazzoni.com.br", dests, "Email Teste", "Segue os arquivos em anexo:", True, "exemplo@mazzoni.com.br", "exemplo123", anexos)
    MsgBox (feitoEnvio)	

A função **sendEmail** fará o envio do email para todos os destinatarios postos no ArrayList. Como pedido, caso seja enviado sem problemas o email retornará verdadeiro, caso não false.
