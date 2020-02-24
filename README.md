# MailSMTP

Projeto de uma Dll criada em C# para fazer envios de email urilizando um servidor SMTP.

----------

## Primeiros passos:

### Gerando um .tlb da dll:

1. Pegue a MailSMTP.dll e copie para a pasta:
	- C:\Windows\System32 (caso seu Windows for x86);
	- C:\Windows\SysWOW64 (caso seu Windows for x64);
2. Abra o Windows PowerShell como Administrador e utilize este comando: cd C:\Windows\Microsoft.NET\Framework\v4.0.30319
3. Execute no Windows PowerShell o seguinte comando:
	- Caso x86: .\RegAsm.exe C:\Windows\System32\MailSMTP.dll /codebase /tlb:C:\Windows\System32\MailSMTP.tlb
	- Caso x64: .\RegAsm.exe C:\Windows\SysWOW64\MailSMTP.dll /codebase /tlb:C:\Windows\SysWOW64\MailSMTP.tlb
	
	**ATENÇÃO:** caso não tenha sido criado o arquivo .tlb voce deve executar novamente o passo 3
	
4. Após a criação do arquivo .tlb voce deve referencia-lo no seu projeto VBA para que seja possivel utilizar suas funcionalidades

----------

## Utilizando o MailSMTP:


### Realizando a refencia no Projeto em VB6:

1. Abra seu projeto VBA e vá ate a aba 'Projeto':

![Screenshot_1](https://user-images.githubusercontent.com/54732019/75123207-339fbd80-5684-11ea-9d79-63d6abe2df59.png)

2. Dentro de 'Projeto' vá até Referências:

![Screenshot_2](https://user-images.githubusercontent.com/54732019/75123213-49ad7e00-5684-11ea-85e8-d803712a1045.png)

3. No formulario de referencias busque pelo arquivo .tlb gerado anteriormente:

![Screenshot_3](https://user-images.githubusercontent.com/54732019/75123219-503bf580-5684-11ea-8ab4-31be1bc74588.png)
	
4. Selecione o arquivo .tlb e esta tudo pronto para começar a utilizar a dll

![Screenshot_4](https://user-images.githubusercontent.com/54732019/75123244-837e8480-5684-11ea-8584-fc71270216e5.png)


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
