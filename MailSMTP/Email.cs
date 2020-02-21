using System;

using System.Collections;
using System.Net.Mail;
using System.Net.Mime;
using System.Net;
using System.Text.RegularExpressions;

namespace MailSMTP
{
    public class Email
    {
        /*
         * mailServer: Servidor de SMTP
         * mailPort: Porta do Servidor
         * fromName: Nome do Remetente
         * tromEmailAddress: Endereço de Email do Remetente
         * toName: Nome do Destinatario
         * toEmailAddress: Endereço de email do Destinatario
         * emailSubject: Assunto do Email
         * emailBody: Corpo do Email
         * attachments: Anexos para o Email
         */
//        public static bool sendEmail(string mailServer, int mailPortServer, string fromName, string fromEmailAddress, string toName,
//          string toEmailAddress, string emailSubject, string emailBody, ArrayList attachments = null, bool reqAuthentication = false,
//          string userNameSLL = "", string passwordSLL = "")
//       {
//            return new sendEmail();
//        }
        public static bool sendEmail(string mailServer, int mailPortServer, string fromName, string fromEmailAddress, string toName,
            string toEmailAddress, string emailSubject, string emailBody, ArrayList attachments = null, bool reqAuthentication = false, 
            string userNameSLL = "", string passwordSLL = "")
        {
            try
            {
                // Valida o Email do Remetente
                bool validaEmailRementente = validaEnderecoEmail(fromEmailAddress);
                if (!validaEmailRementente) return false;
                
                // Valida o Email do Destinatario
                bool validaEmailDestinatario = validaEnderecoEmail(toEmailAddress);
                if (!validaEmailDestinatario) return false;

                // Estrutura o Email para envio
                MailAddress remetente = new MailAddress(fromEmailAddress, fromName);
                MailAddress destinatario = new MailAddress(toEmailAddress, toName);
                MailMessage mensagemEmail = new MailMessage(remetente, destinatario);
                mensagemEmail.Subject = emailSubject;
                mensagemEmail.Body = emailBody;

                // Obtem os anexos contidos em um arquivo arraylist e inclui na mensagem
                if (attachments != null)
                {
                    foreach (string attachment in attachments)
                    {
                        Attachment anexado = new Attachment(attachment, MediaTypeNames.Application.Octet);
                        mensagemEmail.Attachments.Add(anexado);
                    }
                }

                // Faz a conexão com o Client SMTP
                SmtpClient client = new SmtpClient(mailServer, mailPortServer);

                if (reqAuthentication)
                {
                    client.EnableSsl = true;
                    client.UseDefaultCredentials = false;
                    client.Credentials = new NetworkCredential(userNameSLL, passwordSLL);
                }

                client.Send(mensagemEmail);

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        public static bool validaEnderecoEmail(string enderecoEmail)
        {
            try
            {
                //Expressão regular validando o email
                string texto_Validar = enderecoEmail;
                Regex expressaoRegex = new Regex(@"\w+@[a-zA-Z_]+?\.[a-zA-Z]{2,3}");

                // Verifica se o email é valido de acordo com a expressão regular
                if (expressaoRegex.IsMatch(texto_Validar))
                {         
                    return true;
                } 
                else
                {
                    return false;
                }                              
            }
            catch (Exception)
            {
                throw;
            }
        }

    }
}
