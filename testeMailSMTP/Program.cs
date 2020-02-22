
using System;
using System.Collections;
using MailSMTP;

namespace testeMailSMTP
{
    class Program
    {
        static void Main(string[] args)
        {
            ArrayList anexos = new ArrayList();
            anexos.Add(@"D:\35200203929877000150550010000110961533424015.xml");
            ArrayList dest = new ArrayList();
            dest.Add("matheusdiasmazzoni@gmail.com;MatheusDMazzoni");
            dest.Add("matheus.mazzoni@nstecnologia.com.br;NsMazzoni");

            bool t = Email.sendEmail("mail.portalunimec.com.br", 26, "EmpresaTeste", "vendas1@portalunimec.com.br", dest, "TesteDeEmail", "segue em anexo de teste:", true, "vendas1@portalunimec.com.br", "VendMAIL0919!", anexos);
            Console.WriteLine("Foi possivel fazer a comunicação e o envio:" + t);
            Console.ReadLine();
        }
    }
}
