using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Windows.Forms;
using System.Net.Mail;
using System.Net;

namespace Notulensi
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OleDbConnection dbconnection = new OleDbConnection();
            dbconnection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;" +
                @"Data source = C:\Users\Affe Reaflaw\Documents\Visual Studio 2015\Projects\Notulensi\Notulensi\NotulensiDatabase.accdb";
            try
            {
                dbconnection.Open();
                String Tanggal = textBox1.Text.ToString();
                String Bahasan = textBox2.Text.ToString();
                String Presensi = richTextBox1.Text.ToString();
                String Notulensi = richTextBox2.Text.ToString();
                String dbquery = "INSERT into Notulensi(Tanggal, Bahasan, Presensi, Notulensi)values('" + Tanggal + "','" + Bahasan + "', '" + Presensi + "', '" + Notulensi + "')";
                OleDbCommand storedata = new OleDbCommand(dbquery, dbconnection);
                storedata.ExecuteNonQuery();

                MessageBox.Show("Data tersimpan");
                dbconnection.Close();
            }
            catch (Exception error)
            {
                MessageBox.Show("Data tidak tersimpan" + error.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");

                mail.From = new MailAddress(textBox3.Text);
                mail.To.Add(textBox5.Text);
                mail.Subject = textBox1.Text + " " + textBox2.Text;
                mail.Body = "Presensi\n" + richTextBox1.Text + "\nNotulensi\n" + richTextBox2.Text;

                SmtpServer.Port = 587; //587 25 465
                SmtpServer.DeliveryMethod = SmtpDeliveryMethod.Network;
                SmtpServer.UseDefaultCredentials = false;
                SmtpServer.EnableSsl = true;
                SmtpServer.Timeout = 100000;
                SmtpServer.Credentials = new NetworkCredential(textBox3.Text, textBox4.Text);

                SmtpServer.Send(mail);
                MessageBox.Show("Email terkirim");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Email tidak terkirim, cek koneksi anda\n" + ex.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form2 lihatData = new Form2();
            lihatData.Show();
        }
    }
}
