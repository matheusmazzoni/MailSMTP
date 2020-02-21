using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MailSMTP;
namespace testeMailSMTP
{
    class Program
    {
        static void Main(string[] args)
        {
            ArrayList anexos = new ArrayList();
            anexos.Add(@"D:\Freelance Caprind - MatheusMazzoni.pdf");

            bool t = Email.sendEmail("mail.portalunimec.com.br", 26, "EmpresaTeste", "vendas1@portalunimec.com.br", "MatheusMazzoni", "matheusdiasmazzoni@gmail.com", "TesteDeEmail", "segue em anexo de teste:", anexos, true, "vendas1@portalunimec.com.br", "VendMAIL0919!");
            Console.WriteLine(t);
            Console.ReadLine();
        }
    }
}
