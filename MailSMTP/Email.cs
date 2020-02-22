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

        public static bool sendEmail(string mailServer, int mailPortServer, string fromName, string fromEmailAddress, ArrayList recipients,
            string emailSubject, string emailBody, bool reqAuthentication = false, string userNameSLL = "", string passwordSLL = "", ArrayList attachments = null)
        {
            try
            {
                // Cria a mensagem do email
                MailMessage mensagemEmail = new MailMessage();
                mensagemEmail.Subject = emailSubject;
                mensagemEmail.Body = emailBody;
                
                // Valida o Email do Remetente
                if (!validaEnderecoEmail(fromEmailAddress)) return false;                
                mensagemEmail.From = new MailAddress(fromEmailAddress, fromName);

                // Valida o Email do Destinatario
                foreach (string recipient in recipients)
                {
                    String[] recipientSplit = recipient.Split(';');

                    bool validaEmailDestinatario = validaEnderecoEmail(recipientSplit[0]);
                    if (!validaEmailDestinatario) return false;
                    if (recipientSplit.Length > 1)
                    {
                        mensagemEmail.To.Add(new MailAddress(recipientSplit[0].Trim(), recipientSplit[1]));
                    }
                    else
                    {
                        mensagemEmail.To.Add(new MailAddress(recipientSplit[0].Trim()));
                    }
                }

                // Anexar documentos ao corpo do Email
                if (!attachFilesToEmail(mensagemEmail, attachments)) return false;

                return sendMessage(mailServer, mailPortServer, mensagemEmail, reqAuthentication, userNameSLL, passwordSLL);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
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

                if (expressaoRegex.IsMatch(texto_Validar))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {           
                Console.WriteLine(ex);
                throw;
            }
        }
        private static bool attachFilesToEmail(MailMessage mensagemEmail, ArrayList attachments = null)
        {
            try
            {
                if (attachments != null)
                {
                    foreach (string attachment in attachments)
                    {
                        Attachment anexado = new Attachment(attachment, MediaTypeNames.Application.Octet);
                        mensagemEmail.Attachments.Add(anexado);
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return false;
            }
        }
        private static bool sendMessage(string mailServer, int mailPortServer, MailMessage mensagemEmail, bool reqAuthentication = false, string userNameSLL = "", string passwordSLL = "")
        {
            try
            {
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
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return false;
            }
        }

    }
}
