using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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

        public void Math(string temp, int i, int j)
        {
            int t = 0;
            Regex regex = new Regex(@"\w{1}\d{1}");
            if(temp.StartsWith("="))
            {
                if(temp.Length==3)
                {
                    temp = temp.Remove(0,1);
                    MatchCollection matches = regex.Matches(temp);
                    if(matches.Count > 0)
                    {
                        switch (temp[0])
                        {
                            case 'A':
                                t = 0;
                                break;
                            case 'B':
                                t = 1;
                                break;
                            case 'C':
                                t = 2;
                                break;
                            case 'D':
                                t = 3;
                                break;
                            case 'E':
                                t = 4;
                                break;
                            case 'F':
                                t = 5;
                                break;
                            default:
                                break;
                        }
                        //Console.WriteLine( dataGridView1[t,Convert.ToInt32(temp[1])].Value);
                    }
                }
            }
            else if(temp.StartsWith("\'"))
            {
                temp = temp.Trim('\'');
                dataGridView1[i, j].Value = temp;
            }
        }
        public void Import(object i, object j, object data)
        {
            int I = Convert.ToInt32(i);
            int J = Convert.ToInt32(j);
            dataGridView1[I, J].Value = data;
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
                        string sqlExpression = "INSERT INTO dbo.Excel (Row,Col,Data) VALUES ('" + i + "', '" + j + "', '" + dataGridView1[i, j].Value + "')";

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

        private void button1_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Rows.Count; j++)
                {
                    if (dataGridView1[i, j].Value != null)
                    {
                        Math(dataGridView1[i, j].Value.ToString(),i,j);
                    }
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string connect = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\UnbrockenBackBone\source\repos\WindowsFormsApp2\WindowsFormsApp2\Database.mdf;Integrated Security=True";

            string sqlExpression = "SELECT * FROM Excel";
            using (SqlConnection connection = new SqlConnection(connect))
            {
                connection.Open();
                SqlCommand command = new SqlCommand(sqlExpression, connection);
                SqlDataReader reader = command.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        object row = reader.GetValue(1);
                        object col = reader.GetValue(2);
                        object data = reader.GetValue(3);
                        Import(row, col, data);
                    }
                }

                reader.Close();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns.Clear();
            Start();
        }
    }
}
