using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication2
{
    public partial class Form1 : Form
    {
        OleDbConnection connection;
        SqlConnection server_connection;
        DataTable table;
        string sql_server_string;
        string server, username, password, database;
        string file;

        public Form1()
        {
            InitializeComponent();
            table = new DataTable();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string connection_string;

            openFileDialog1.Filter = "mdb files(*.mdb)|*.mdb";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.ShowDialog();

            file = openFileDialog1.FileName;
            textBox1.Text = file;

            connection_string = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + file + "; Jet OLEDB:Database Password=0";
            connection = new OleDbConnection(connection_string);
            connection.Open();

            textBox5.Text = connection_string;
            label7.Text = "Access";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string query = richTextBox1.Text;

            if (label7.Text == "SQL Server")
            {
                SqlDataAdapter adapterSql = new SqlDataAdapter(query, server_connection);

                table.Clear();
                adapterSql.Fill(table);
                dataGridView1.DataSource = table;
            }
            else if (label7.Text == "Access")
            {
                    OleDbDataAdapter adapterOleDb = new OleDbDataAdapter(query, connection);

                    table.Clear();
                    adapterOleDb.Fill(table);
                    dataGridView1.DataSource = table;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string query = richTextBox1.Text;

            if (label7.Text == "SQL Server")
            {
                SqlDataAdapter adapterSql = new SqlDataAdapter(query, server_connection);
                SqlCommandBuilder commandBuilder = new SqlCommandBuilder(adapterSql);

                adapterSql.Update(table);
                table.Clear();
                adapterSql.Fill(table);
                dataGridView1.DataSource = table;
            }
            else if (label7.Text == "Access")
            {
                OleDbDataAdapter adapterOleDb = new OleDbDataAdapter(query, connection);
                OleDbCommandBuilder cb = new OleDbCommandBuilder(adapterOleDb);

                cb.GetDeleteCommand();
                adapterOleDb.Update(table);
                table.Clear();
                adapterOleDb.Fill(table);
                dataGridView1.DataSource = table;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            server = textBox2.Text;
            database = textBox3.Text;
            username = textBox6.Text;
            password = textBox4.Text;
            sql_server_string = "Server=" + server + ";Database=" + database + ";Trusted_Connection=True;";
            server_connection = new SqlConnection(sql_server_string);
            server_connection.Open();

            textBox5.Text = sql_server_string;
            label7.Text = "SQL Server";
        }
    }
}

