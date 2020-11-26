using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Mail;
using System.Text;

namespace SendMail
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Write("Enviar emails(s/n)?: ");
            string continuar = Console.ReadLine();

            if (!continuar.ToLower().Equals("s"))
            {
                return;
            }

            const string PATH = @"C:\Users\Tacio\Desktop\Emails.csv";

            using (var reader = new StreamReader(PATH, Encoding.GetEncoding("iso-8859-1")))
            {
                List<CSV> escolas = new List<CSV>();

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(';');

                    escolas.Add(new CSV
                    {
                        Escola = values[0],
                        EscolaEmail = values[1],
                        GestorEmail = values[2],
                        SecretarioEmail = values[3]
                    });
                }

                foreach (CSV escola in escolas)
                {
                    SendMail(escola);
                }

            }
        }

        static void SendMail(CSV escola)
        {
            // Constantes do programa
            const string Mensal = "Mensal";
            const string Bimestral = "Bimestral";
            //const string NOTA10COMPARATIVO = "NOTA 10 Comparativo";

            // Variaveis de configuração do email
            const int PORT = 25;
            const string EMAIL_SIEM = "siem_seduc@hotmail.com";
            const string SENHA_SIEM = "3dUc@c@02";
            const string SUBJECT = "Relatório Educação Nota 10 2ª Unidade";
            string BODY = $@"<html>
<head></head>
<body style='font-family: Calibri,Arial,Helvetica,sans-serif;'>
<p>Boa Noite!!</p>
<p>Segue em anexo o relatório do Educação Nota 10 - 2ª Unidade Mensal(Caso possua Ensino Fundamental I) e Bimestral(Ensino Fundamental II, Educação Infantil e EJA) da escola {escola.Escola}.</p>
<br />
<div style='text-align: center;  letter-spacing: .3px'>
<p style='margin: 0px; font-size: 14; font-weight: 600'>Atenciosamente,
<br />
<br />
<p style='margin: 0px;'>Equipe SIEM</p>
<p style='margin: 0px;'>Setor de Informações da Educação Municipal</p>
<p style='margin: 0px;'>Superintendência de Gestão - SEDUC - Juazeiro / BA</p>
<p style='margin: 0px;'>(74) 3612-3379</p>
</div>";

            // Variaveis de teste de anexos
            const bool hasMensalAttachment = true;
            const bool hasBimestralAttachment = true;
            //const bool hasNOTA10ComparativoAttachment = false;


            using (MailMessage mail = new MailMessage())
            using (SmtpClient smtp = new SmtpClient("smtp.live.com"))
            {

                mail.From = new MailAddress(EMAIL_SIEM);
                mail.Subject = SUBJECT;
                mail.Body = BODY;
                mail.IsBodyHtml = true;


                AddMailAddresses(escola, mail);


                if(hasMensalAttachment || hasBimestralAttachment)
                {
                    if (hasMensalAttachment)
                    {
                        string attachmentPath = $@"C:\Users\Tacio\Desktop\{Mensal}\{escola.Escola} - Mensal.pdf";
                        AddMailAttachment(mail, attachmentPath);
                    }

                    if (hasBimestralAttachment)
                    {
                        string attachmentPath = $@"C:\Users\Tacio\Desktop\{Bimestral}\{escola.Escola}.pdf";
                        AddMailAttachment(mail, attachmentPath);
                    }

                }                

                smtp.Port = PORT;
                smtp.Credentials = new System.Net.NetworkCredential(EMAIL_SIEM, SENHA_SIEM);
                smtp.EnableSsl = true;
                

                try
                {
                    smtp.Send(mail);
                    Console.WriteLine($"Email para a escola: {escola.Escola} enviado com sucesso!");
                } catch {
                    Console.WriteLine($"Email para a escola: {escola.Escola} falhou!");
                }
            }
        }

        private static void AddMailAttachment(MailMessage mail, string attachmentPath)
        {
            try
            {

                Attachment attachment = new Attachment(attachmentPath);
                if (attachment != null)
                {
                    mail.Attachments.Add(attachment);
                }

            } catch { }
        }

        private static void AddMailAddresses(CSV escola, MailMessage mail)
        {
            if (!String.IsNullOrEmpty(escola.EscolaEmail))
            {
                mail.To.Add(escola.EscolaEmail);
            }
            if (!String.IsNullOrEmpty(escola.GestorEmail))
            {
                mail.To.Add(escola.GestorEmail);
            }
            if (!String.IsNullOrEmpty(escola.SecretarioEmail))
            {
                mail.To.Add(escola.SecretarioEmail);
            }
        }
    }
}
