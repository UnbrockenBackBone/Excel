using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        DataGridViewTextBoxColumn[] column = new DataGridViewTextBoxColumn[5];
        public Form1()
        {
            InitializeComponent();
            Start();
        }
        public void Start()
        {
            int m = 10;
            int temp = 1;


            for (int i = 0; i < column.Length; i++)
            {
                column[i] = new DataGridViewTextBoxColumn();
                column[i].HeaderText = "" + temp;
                column[i].Name = "" + temp;
                temp++;
            }

            this.dataGridView1.Columns.AddRange(column);
            for (int i = 0; i < m; i++)
            {
                this.dataGridView1.Rows.Add();
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            string connect = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\UnbrockenBackBone\source\repos\WindowsFormsApp2\WindowsFormsApp2\Database.mdf;Integrated Security=True";

            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Rows.Count; j++)
                {
                    if (dataGridView1[i, j].Value != null)
                    {
                        string sqlExpression = "INSERT INTO dbo.Excel (Row,Col,Data) VALUES ('" + dataGridView1.Columns[i] + "', '" + dataGridView1.Rows[j] + "', '" + dataGridView1[i, j].Value + "')";

                        using (SqlConnection connection = new SqlConnection(connect))
                        {
                            connection.Open();
                            SqlCommand command = new SqlCommand(sqlExpression, connection);
                            try
                            {
                                command.Connection = connection;
                                command.ExecuteNonQuery();
                            }
                            catch (SqlException a)
                            {
                                Console.WriteLine(a.ToString());
                            }
                        }

                    }
                }
            }
        }
    }
}
