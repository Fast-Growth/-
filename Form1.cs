using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;

namespace Расписание_занятий
{
    public partial class Form1 : Form
    {
        private string sqlText;
        private SqlDataAdapter adapter;
        private DataSet dataSet;
        private DialogResult result;
        public Form1()
        {
            InitializeComponent();
        }

        private void LoadData()
        {
            try
            {              
                this.неделиTableAdapter.Fill(this.расписание_занятийDataSet.Недели);               
                this.графики_обученияTableAdapter.Fill(this.расписание_занятийDataSet.Графики_обучения);             
                this.таблица_соединенийTableAdapter.Fill(this.расписание_занятийDataSet.Таблица_соединений);                
                this.дисциплиныTableAdapter.Fill(this.расписание_занятийDataSet.Дисциплины);
                this.классыTableAdapter.Fill(this.расписание_занятийDataSet.Классы);
                sqlText = "select [Неделя], [Класс], [Дисциплина] from [Классы], [Дисциплины], [Таблица_соединений], [Графики_обучения], [Недели] " +
                    "Where [Классы].[ID_класса] = [Таблица_соединений].[ID_класса] and " +
                    "[Дисциплины].[ID_дисциплины] = [Таблица_соединений].[ID_дисциплины] and " +
                    "[Недели].[ID_недели] = [Графики_обучения].[ID_недели] and " +
                    "[Таблица_соединений].[ID_соединения] = [Графики_обучения].[ID_соединения] and " +
                    "[Недели].[Неделя] = '" + comboBox1.Text + "'" +
                    "ORDER BY [Классы].[Класс] asc"; 
                adapter = new SqlDataAdapter(sqlText, Class.DataBase.connStr);
                dataSet = new DataSet();
                adapter.Fill(dataSet, "Классы, Дисциплины, Таблица_соединений, Графики_обучения, Недели");
                dataGridView1.DataSource = dataSet.Tables["Классы, Дисциплины, Таблица_соединений, Графики_обучения, Недели"];
                Load_zan();
                Load_dis();
            }
            catch
            {
                MessageBox.Show("Проблема с БД!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void Load_zan()
        {
            sqlText = "SELECT [Дисциплина] FROM [Классы], [Дисциплины], [Таблица_соединений], [Графики_обучения], [Недели]" +
                "WHERE[Классы].[ID_класса] = [Таблица_соединений].[ID_класса] AND" +
                "[Дисциплины].[ID_дисциплины] = [Таблица_соединений].[ID_дисциплины] AND" +
                "[Недели].[ID_недели] = [Графики_обучения].[ID_недели] AND" +
                "[Таблица_соединений].[ID_соединения] = [Графики_обучения].[ID_соединения] AND" +
                "[Классы].[Класс] = '" + comboBox4.Text + "' AND" +
                "[Недели].[Неделя] = '"+comboBox3.Text+"'";
            adapter = new SqlDataAdapter(sqlText, Class.DataBase.connStr);
            DataTable table = new DataTable();
            adapter.Fill(table);
            comboBox5.DataSource = table;
            comboBox5.DisplayMember = "Дисциплина";
        }

        private void Load_dis()
        {
            sqlText = "SELECT [Дисциплина] FROM [Классы], [Дисциплины], [Таблица_соединений]" +
                "WHERE[Классы].[ID_класса] = [Таблица_соединений].[ID_класса] AND" +
                "[Дисциплины].[ID_дисциплины] = [Таблица_соединений].[ID_дисциплины] AND" +
                "[Классы].[Класс] = '" + comboBox4.Text + "'";
            adapter = new SqlDataAdapter(sqlText, Class.DataBase.connStr);
            DataTable table = new DataTable();
            adapter.Fill(table);
            comboBox8.DataSource = table;
            comboBox8.DisplayMember = "Дисциплина";
        }

        private void Combick()
        {
            sqlText = "SELECT Дисциплина FROM Дисциплины";
            adapter = new SqlDataAdapter(sqlText, Class.DataBase.connStr);
            DataTable table = new DataTable();
            adapter.Fill(table);
            comboBox6.DataSource = table;
            comboBox6.DisplayMember = "Дисциплина";
            comboBox6.ValueMember = "id_дисциплины";
        }

        private void ReloadData()
        {
            try
            {
                LoadData();
            }
            catch
            {
                MessageBox.Show("Проблема с БД!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadData();
        }

        private void btnMenuStrip_Click(object sender, EventArgs e)
        {
            string Text = comboBox1.Text;
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workBook = excelApp.Workbooks.Add();
            Excel.Worksheet workSheet = workBook.Sheets[1];
            workSheet.Name = "Отчёт";
            workSheet.Cells[2, 3] = "Расписание занятий в "+ comboBox3.Text +"";
            Excel.Range rng1 = workSheet.Range[workSheet.Cells[2, 3], workSheet.Cells[2, 3]];
            rng1.Cells.Font.Name = "Times New Roman";
            rng1.Cells.Font.Size = 24;
            rng1.Font.Bold = true;
            rng1.Cells.Font.Color = ColorTranslator.ToOle(Color.Green);
            int b = 5;
            for (int g = 2; g <= 8; g++)
            {
                workSheet.Cells[4, g] = ""+b+" класс";
                workSheet.Columns[g].ColumnWidth = 20;
                Excel.Range rng2 = workSheet.Range[workSheet.Cells[4, g], workSheet.Cells[4, g]];
                rng2.Font.Bold = true;
                
                string SqlText = "select [Дисциплины].[Дисциплина] from [Классы], [Дисциплины], [Таблица_соединений], [Графики_обучения], [Недели] " +
                    "Where [Классы].[ID_класса] = [Таблица_соединений].[ID_класса] and " +
                    "[Дисциплины].[ID_дисциплины] = [Таблица_соединений].[ID_дисциплины] and " +
                    "[Недели].[ID_недели] = [Графики_обучения].[ID_недели] and " +
                    "[Таблица_соединений].[ID_соединения] = [Графики_обучения].[ID_соединения] and " +
                    "[Недели].[Неделя] = '"+Text+"' and " +
                    "[Классы].[Класс] = '" + b + "'";
                b++;
                SqlDataAdapter adapter = new SqlDataAdapter(SqlText, Class.DataBase.connStr);
                DataTable table = new DataTable();
                adapter.Fill(table);
                int i = 5;
                foreach (DataRow row in table.Rows)
                {
                    workSheet.Cells[i, g] = row["Дисциплина"];
                    i++;
                    Excel.Range rng3 = workSheet.Range[workSheet.Cells[4, g], workSheet.Cells[i - 1, g]];
                    rng3.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
                    rng3.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
                    rng3.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous;
                    rng3.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous;
                    rng3.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;
                    rng3.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
                }
            }
            excelApp.Visible = true;
            excelApp.UserControl = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataSet.Tables["Классы, Дисциплины, Таблица_соединений, Графики_обучения, Недели"].Clear();
            sqlText = "select Неделя, Класс, Дисциплина from Классы, Дисциплины, Таблица_соединений, Графики_обучения, Недели " +
                    "Where Классы.ID_класса = Таблица_соединений.ID_класса and " +
                    "Дисциплины.ID_дисциплины = Таблица_соединений.ID_дисциплины and " +
                    "Недели.ID_недели = Графики_обучения.ID_недели and " +
                    "Таблица_соединений.ID_соединения = Графики_обучения.ID_соединения and " +
                    "[Недели].[Неделя] = '" + comboBox1.Text + "'" +
                    "ORDER BY [Классы].[Класс] asc";
            adapter = new SqlDataAdapter(sqlText, Class.DataBase.connStr);
            dataSet = new DataSet();
            adapter.Fill(dataSet, "Классы, Дисциплины, Таблица_соединений, Графики_обучения, Недели");
            dataGridView1.DataSource = dataSet.Tables["Классы, Дисциплины, Таблица_соединений, Графики_обучения, Недели"];
        }

        private void button1_Click(object sender, EventArgs e)
        {
            sqlText = "select Дисциплины.Дисциплина as '" + comboBox2.Text + " класс' from Классы, Дисциплины, Таблица_соединений, Графики_обучения, Недели " +
                    "Where Классы.ID_класса = Таблица_соединений.ID_класса and " +
                    "Дисциплины.ID_дисциплины = Таблица_соединений.ID_дисциплины and " +
                    "Недели.ID_недели = Графики_обучения.ID_недели and " +
                    "Таблица_соединений.ID_соединения = Графики_обучения.ID_соединения and " +
                    "Недели.Неделя = '" + comboBox1.Text + "' and " +
                    "Классы.Класс = '" + comboBox2.Text + "'";
            adapter = new SqlDataAdapter(sqlText, Class.DataBase.connStr);
            dataSet = new DataSet();
            adapter.Fill(dataSet, "Классы, Дисциплины, Таблица_соединений, Графики_обучения, Недели");
            dataGridView1.DataSource = dataSet.Tables["Классы, Дисциплины, Таблица_соединений, Графики_обучения, Недели"];
        }

        private void button5_Click(object sender, EventArgs e)
        {
            sqlText = "SELECT [Класс], [Дисциплина] FROM [Классы], [Дисциплины], [Таблица_соединений]" +
                "WHERE[Классы].[ID_класса] = [Таблица_соединений].[ID_класса] AND" +
                "[Дисциплины].[ID_дисциплины] = [Таблица_соединений].[ID_дисциплины] AND" +
                "[Классы].[Класс] = '" + comboBox2.Text + "'"; 
            adapter = new SqlDataAdapter(sqlText, Class.DataBase.connStr);
            dataSet = new DataSet();
            adapter.Fill(dataSet, "Классы, Дисциплины, Таблица_соединений, Графики_обучения, Недели");
            dataGridView1.DataSource = dataSet.Tables["Классы, Дисциплины, Таблица_соединений, Графики_обучения, Недели"];
        }

        private void button7_Click(object sender, EventArgs e)
        {
            result = MessageBox.Show("Удалить запись " + comboBox6.Text + "?", "Внимание", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (result == DialogResult.OK) 
            {
                sqlText = "DELETE FROM [Дисциплины] WHERE [Дисциплина] =  '" + comboBox6.Text + "'";
                SqlConnection connStr = new SqlConnection(Class.DataBase.connStr);
                connStr.Open();
                SqlCommand cmd = new SqlCommand(sqlText, connStr);
                cmd.ExecuteNonQuery();
                connStr.Close();
                Combick();
                MessageBox.Show("Действие успешно выполнено", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Действие отменено", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void Button6_Click(object sender, EventArgs e)
        {
            result = MessageBox.Show("Добавить запись " + textBox1.Text + "?", "Внимание", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (result == DialogResult.OK)
            {
                if (textBox1.Text != "")
                {
                    sqlText = "INSERT INTO [Дисциплины] ([Дисциплина])" +
                    "VALUES ('" + textBox1.Text + "')";
                    SqlConnection connStr = new SqlConnection(Class.DataBase.connStr);
                    connStr.Open();
                    SqlCommand cmd = new SqlCommand(sqlText, connStr);
                    cmd.ExecuteNonQuery();
                    connStr.Close();
                    Combick();
                    MessageBox.Show("Действие успешно выполнено", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Действие отменено, так как текстовое поле пустое", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            result = MessageBox.Show("Изменить запись " + comboBox6.Text + " на " + textBox1.Text + "?", "Внимание", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (result == DialogResult.OK)
            {
                sqlText = "UPDATE [Дисциплины]" +
                "SET [Дисциплина] = '" + textBox1.Text + "'" +
                "WHERE [Дисциплина] = '" + comboBox6.Text + "'";
                SqlConnection connStr = new SqlConnection(Class.DataBase.connStr);
                connStr.Open();
                SqlCommand cmd = new SqlCommand(sqlText, connStr);
                cmd.ExecuteNonQuery();
                connStr.Close();
                Combick();
                MessageBox.Show("Действие успешно выполнено", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Действие отменено", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            result = MessageBox.Show("Добавить занятие " + comboBox6.Text + " в " + comboBox7.Text + " класс?", "Внимание", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (result == DialogResult.OK)
            {
                sqlText = "INSERT INTO [Таблица_соединений] ([ID_класса], [ID_дисциплины])" +
                "VALUES ('" + comboBox7.SelectedValue + "','" + comboBox6.SelectedValue + "')";
                SqlConnection connStr = new SqlConnection(Class.DataBase.connStr);
                connStr.Open();
                SqlCommand cmd = new SqlCommand(sqlText, connStr);
                cmd.ExecuteNonQuery();
                connStr.Close();
                Load_dis();
                MessageBox.Show("Действие успешно выполнено", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Действие отменено", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            result = MessageBox.Show("Удалить занятие " + comboBox6.Text + " в " + comboBox7.Text + " классе?", "Внимание", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (result == DialogResult.OK)
            {
                sqlText = "DELETE FROM [Таблица_соединений]" +
                "WHERE [ID_класса] = '" + comboBox7.SelectedValue + "' AND [ID_дисциплины] = '" + comboBox6.SelectedValue + "'";
                SqlConnection connStr = new SqlConnection(Class.DataBase.connStr);
                connStr.Open();
                SqlCommand cmd = new SqlCommand(sqlText, connStr);
                cmd.ExecuteNonQuery();
                connStr.Close();
                Load_dis();
                MessageBox.Show("Действие успешно выполнено", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Действие отменено", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void обновитьБДToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ReloadData();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            result = MessageBox.Show("Добавить дисциплину " + comboBox8.Text + " в " + comboBox3.Text + " для " + comboBox4.Text + " класса?", "Внимание", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (result == DialogResult.OK)
            {
                sqlText = "INSERT INTO [Графики_обучения]([ID_недели], [ID_соединения])" +
                "SELECT [Недели].[ID_недели], [Таблица_соединений].[ID_соединения]" +
                "FROM [Недели], [Таблица_соединений], [Классы], [Дисциплины]" +
                "WHERE[Классы].[ID_класса] = [Таблица_соединений].[ID_класса] AND" +
                "[Дисциплины].[ID_дисциплины] = [Таблица_соединений].[ID_дисциплины] AND" +
                "[Классы].[Класс] = '" + comboBox4.Text + "' AND" +
                "[Дисциплины].[Дисциплина] = '" + comboBox8.Text + "' AND" +
                "[Недели].[Неделя] = '" + comboBox3.Text + "'";
                SqlConnection connStr = new SqlConnection(Class.DataBase.connStr);
                connStr.Open();
                SqlCommand cmd = new SqlCommand(sqlText, connStr);
                cmd.ExecuteNonQuery();
                connStr.Close();
                Load_zan();
                MessageBox.Show("Действие успешно выполнено", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Действие отменено", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            result = MessageBox.Show("Удалить занятие " + comboBox5.Text + " в " + comboBox3.Text + " для " + comboBox4.Text + " класса?", "Внимание", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (result == DialogResult.OK)
            {
                sqlText = "DELETE FROM [Графики_обучения] WHERE" +
                "[ID_соединения] = (SELECT[ID_соединения] FROM[Таблица_соединений], [Классы], [Дисциплины]" +
                "WHERE[Классы].[ID_класса] = [Таблица_соединений].[ID_класса] AND" +
                "[Дисциплины].[ID_дисциплины] = [Таблица_соединений].[ID_дисциплины] AND" +
                "[Классы].[Класс] = '" + comboBox4.Text + "' AND" +
                "[Дисциплины].[Дисциплина] = '" + comboBox5.Text + "') AND" +
                "[ID_недели] = (SELECT[ID_недели] FROM[Недели]" +
                "WHERE[Неделя] = '" + comboBox3.Text + "')";
                SqlConnection connStr = new SqlConnection(Class.DataBase.connStr);
                connStr.Open();
                SqlCommand cmd = new SqlCommand(sqlText, connStr);
                cmd.ExecuteNonQuery();
                connStr.Close();
                Load_zan();
                MessageBox.Show("Действие успешно выполнено", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Действие отменено", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            result = MessageBox.Show("Изменить занятие " + comboBox5.Text + " на " + comboBox8.Text + " для " + comboBox4.Text + " класса?", "Внимание", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (result == DialogResult.OK)
            {
                sqlText = "UPDATE [Графики_обучения]" +
                "SET[ID_соединения] = (" +
                "SELECT[ID_соединения] FROM[Таблица_соединений], [Классы], [Дисциплины]" +
                "WHERE[Классы].[ID_класса] = [Таблица_соединений].[ID_класса] AND" +
                "[Дисциплины].[ID_дисциплины] = [Таблица_соединений].[ID_дисциплины] AND" +
                "[Классы].[Класс] = '" + comboBox4.Text + "' AND" +
                "[Дисциплины].[Дисциплина] = '" + comboBox8.Text + "') " +
                "WHERE[ID_соединения] = (" +
                "SELECT[ID_соединения] FROM[Таблица_соединений], [Классы], [Дисциплины]" +
                "WHERE[Классы].[ID_класса] = [Таблица_соединений].[ID_класса] AND" +
                "[Дисциплины].[ID_дисциплины] = [Таблица_соединений].[ID_дисциплины] AND" +
                "[Классы].[Класс] = '" + comboBox4.Text + "' AND" +
                "[Дисциплины].[Дисциплина] = '" + comboBox5.Text + "')";
                SqlConnection connStr = new SqlConnection(Class.DataBase.connStr);
                connStr.Open();
                SqlCommand cmd = new SqlCommand(sqlText, connStr);
                cmd.ExecuteNonQuery();
                connStr.Close();
                Load_zan();
                MessageBox.Show("Действие успешно выполнено", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Действие отменено", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            sqlText = "SELECT [Дисциплина] FROM [Классы], [Дисциплины], [Таблица_соединений], [Графики_обучения], [Недели]" +
                "WHERE[Классы].[ID_класса] = [Таблица_соединений].[ID_класса] AND" +
                "[Дисциплины].[ID_дисциплины] = [Таблица_соединений].[ID_дисциплины] AND" +
                "[Недели].[ID_недели] = [Графики_обучения].[ID_недели] AND" +
                "[Таблица_соединений].[ID_соединения] = [Графики_обучения].[ID_соединения] AND" +
                "[Классы].[Класс] = '"+ comboBox4.Text + "' AND" +
                "[Недели].[Неделя] = '"+comboBox3.Text+"'";
            adapter = new SqlDataAdapter(sqlText, Class.DataBase.connStr);
            DataTable table = new DataTable();
            adapter.Fill(table);
            comboBox5.DataSource = table;
            comboBox5.DisplayMember = "Дисциплина";
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            sqlText = "SELECT [Дисциплина] FROM [Классы], [Дисциплины], [Таблица_соединений], [Графики_обучения], [Недели]" +
                "WHERE[Классы].[ID_класса] = [Таблица_соединений].[ID_класса] AND" +
                "[Дисциплины].[ID_дисциплины] = [Таблица_соединений].[ID_дисциплины] AND" +
                "[Недели].[ID_недели] = [Графики_обучения].[ID_недели] AND" +
                "[Таблица_соединений].[ID_соединения] = [Графики_обучения].[ID_соединения] AND" +
                "[Классы].[Класс] = '" + comboBox4.Text + "' AND" +
                "[Недели].[Неделя] = '" + comboBox3.Text + "'";
            adapter = new SqlDataAdapter(sqlText, Class.DataBase.connStr);
            DataTable table = new DataTable();
            adapter.Fill(table);
            comboBox5.DataSource = table;
            comboBox5.DisplayMember = "Дисциплина";
        }
    }
}
