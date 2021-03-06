﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Windows.Forms;

namespace Notulensi
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            OleDbConnection dbconnection = new OleDbConnection();
            dbconnection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;" +
                @"Data source = C:\Users\Affe Reaflaw\Documents\Visual Studio 2015\Projects\Notulensi\Notulensi\NotulensiDatabase.accdb";
            try
            {
                dbconnection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = dbconnection;
                string myQuery = "SELECT TOP 100 * FROM Notulensi ORDER BY ID desc";
                command.CommandText = myQuery;
                OleDbDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    listBox1.Items.Add(reader["Bahasan"].ToString());
                }
                dbconnection.Close();
            }
            catch (Exception error)
            {
                MessageBox.Show("Data Notulensi tidak ada" + error.Message);
            }
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

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            OleDbConnection dbconnection = new OleDbConnection();
            dbconnection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;" +
                @"Data source = C:\Users\Affe Reaflaw\Documents\Visual Studio 2015\Projects\Notulensi\Notulensi\NotulensiDatabase.accdb";
            try
            {
                dbconnection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = dbconnection;
                string myQuery = "SELECT * FROM Notulensi where Bahasan = '" + listBox1.Text+"'";
                command.CommandText = myQuery;
                OleDbDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    textBox1.Text=(reader["Tanggal"].ToString());
                    textBox2.Text = (reader["Bahasan"].ToString());
                    richTextBox1.Text = (reader["Presensi"].ToString());
                    richTextBox2.Text = (reader["Notulensi"].ToString());
                }
                dbconnection.Close();
            }
            catch (Exception error)
            {
                MessageBox.Show("Data Notulensi tidak ada" + error.Message);
            }
        }
    }
}
