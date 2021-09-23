using System;
using System.Configuration;
using System.Net;
using System.Windows.Forms;
using System.Net.Mail;

namespace SendEmail
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            var emailFrom = ConfigurationManager.AppSettings["EmailFrom"];
            var team = ConfigurationManager.AppSettings["Team"];
            var subject = ConfigurationManager.AppSettings["Subject"];
            txbFrom.Text = emailFrom;
            txbTo.Text = team;
            txbSubject.Text = subject;
        }

        private void btnSend_Click(object sender, EventArgs e)
        {
            EnviarEmail();
        }

        private void EnviarEmail()
        {
            var smtpCredential = ConfigurationManager.AppSettings["Smtp"];
            var user = ConfigurationManager.AppSettings["User"];
            var password = ConfigurationManager.AppSettings["Password"];
            var emailFrom = ConfigurationManager.AppSettings["EmailFrom"];
            var emailTo = ConfigurationManager.AppSettings["ListEmails"];

            try
            {
                using (SmtpClient smtp = new SmtpClient())
                {
                    using (MailMessage email = new MailMessage())
                    {
                        //Server SMTP
                        smtp.Host = smtpCredential;
                        smtp.UseDefaultCredentials = false;
                        smtp.Credentials = new NetworkCredential(user, password);
                        smtp.Port = 587;
                        smtp.EnableSsl = false;

                        //Email
                        email.From = new MailAddress(emailFrom);
                        foreach (var address in emailTo.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
                        {
                            email.To.Add(address);
                        }

                        email.Subject = txbSubject.Text;
                        email.IsBodyHtml = false;
                        email.Body = txbMessage.Text;

                        //Attachs
                        if (lblAttach.Text != "")
                        {
                            var attach = lblAttach.Text.ToString().Split(';');
                            for (int i = 0; i < attach.Length; i++)
                                email.Attachments.Add(new Attachment(attach[i]));
                        }

                        //Send Email
                        smtp.Send(email);
                    }
                }

                MessageBox.Show("Email Enviado!");
            }
            catch(Exception ex)
            {
                MessageBox.Show("Erro: " + ex.Message);
            }
        }

        private void lblAttach_Click(object sender, EventArgs e)
        {
            var attach = new OpenFileDialog();

            attach.Multiselect = true;
            attach.Title = "Attach files";
            if (attach.ShowDialog() == DialogResult.OK)
                for (int i = 0; i < attach.FileNames.Length;i++)
                    if (i == 0)
                        lblAttach.Text = attach.FileNames[i];
                    else
                        lblAttach.Text = lblAttach.Text + "; " + attach.FileNames[i];
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            txbMessage.Clear();
            txbSubject.Clear();
        }
    }
}
