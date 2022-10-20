using System.Data;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using org.mariuszgromada.math.mxparser;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Excel_Parody
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        DataTable table = new DataTable();
        Dictionary<string, string[]> Cell_Dictionary = new Dictionary<string, string[]>();
        private void Form1_Load(object sender, EventArgs e)
        {
            for (int i = 1; i < 7; i++)
            {
                //dataGridView1.Columns.Add(Convert.ToString(i), Convert.ToString(i));
                table.Columns.Add((table.Columns.Count + 1).ToString());
            }
            for (int i = 1; i < 7; i++)
            {
                table.Rows.Add();
            }
            for (int i = 1; i < 7; i++)
            {
                for(int a = 1; a < 7; a++)
                {
                    Cell_Dictionary.Add("R" + i.ToString() + "C" + a.ToString(), new []{"",""});
                }
            }

            dataGridView1.DataSource = table;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        // Add Row
        private void button1_Click(object sender, EventArgs e)
        {
            table.Rows.Add();
            int count_of_rows = table.Rows.Count;
            int count_of_columns = table.Columns.Count;
            for (int a = 1; a <= count_of_columns; a++)
            {
                Cell_Dictionary.Add("R" + (count_of_rows).ToString() + "C" +  a.ToString(), new[] {"",""});
            }
        }
        // Add Column
        private void button2_Click(object sender, EventArgs e)
        {
            //dataGridView1.Columns.Add(Convert.ToString(dataGridView1.Columns.Count + 1), Convert.ToString(dataGridView1.Columns.Count + 1));
            table.Columns.Add((table.Columns.Count + 1).ToString());
            int count_of_rows = table.Rows.Count;
            int count_of_columns = table.Columns.Count;
            for (int a = 1; a <= count_of_rows; a++)
            {
                Cell_Dictionary.Add("R" + (a).ToString() + "C" + (count_of_columns).ToString(), new[] {"",""});
            }
        }
        // Adding Indexes to Rows
        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var grid = sender as DataGridView;
            var rowIdx = (e.RowIndex + 1).ToString();

            var centerformat = new StringFormat()
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center,
            };
            var headerBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth, e.RowBounds.Height);
            e.Graphics.DrawString(rowIdx, this.Font, SystemBrushes.ControlText, headerBounds, centerformat);

        }
        // Delete Column
        private void button2_Click_1(object sender, EventArgs e)
        {
            int column_ind = dataGridView1.CurrentCell.ColumnIndex;
            int count_of_rows = table.Rows.Count;
            int count_of_columns = table.Columns.Count;
            table.Columns.RemoveAt(column_ind);
            for (int a = 1; a <= count_of_rows; a++)
            {
                for (int i = column_ind + 1; i < count_of_columns; i++)
                {
                    Cell_Dictionary["R" + (a.ToString()) + "C" + (i.ToString())][0] = Cell_Dictionary["R" + (a.ToString()) + "C" + (i + 1).ToString()][0];
                    Cell_Dictionary["R" + (a.ToString()) + "C" + (i.ToString())][1] = Cell_Dictionary["R" + (a.ToString()) + "C" + (i + 1).ToString()][1];
                }
           
            }
            for (int i = 1; i <= count_of_rows; i++)
            {
                Cell_Dictionary.Remove("R" + (i).ToString() + "C" + (count_of_columns).ToString());
            }

            if (table.Columns.Count != column_ind)

            {                                
                for (int i = column_ind; i < table.Columns.Count; i++)
                {
                    int value_int = Int32.Parse(table.Columns[i].ColumnName);
                    table.Columns[i].ColumnName = Convert.ToString(value_int - 1);
                }
            }
            else
            {

            }
        }

        // Delete Row
        private void button1_Click_1(object sender, EventArgs e)
        {
            int rows_ind = dataGridView1.CurrentCell.RowIndex;
            int count_of_rows = table.Rows.Count;
            int count_of_columns = table.Columns.Count;
            table.Rows.RemoveAt(rows_ind);
            for (int a = 1; a <= count_of_columns; a++)
            {
                for (int i = rows_ind + 1; i < count_of_rows; i++)
                {
                    Cell_Dictionary["R" + (i.ToString()) + "C" + (a.ToString())][0] = Cell_Dictionary["R" + (i+1).ToString() + "C" + (a.ToString())][0];
                    Cell_Dictionary["R" + (i.ToString()) + "C" + (a.ToString())][1] = Cell_Dictionary["R" + (i + 1).ToString() + "C" + (a.ToString())][1];
                }

            }
            for (int i = 1; i <= count_of_columns; i++)
            {
                Cell_Dictionary.Remove("R" + (count_of_rows).ToString() + "C" + (i.ToString()));
            }


        }
        // Show Cell
        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
        // Show Expression
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int rows_index = e.RowIndex;
            int columns_index = e.ColumnIndex;
            if (rows_index >= 0 && columns_index >= 0)
            {
                if (Cell_Dictionary["R" + (rows_index + 1).ToString() + "C" + (columns_index + 1).ToString()][1] != "")
                {
                    textBox1.Text = Cell_Dictionary["R" + (rows_index + 1).ToString() + "C" + (columns_index + 1).ToString()][0];
                    textBox2.Text = "R" + (rows_index + 1).ToString() + "C" + (columns_index + 1).ToString();
                }
                else
                {
                    textBox1.Text = "";
                    textBox2.Text = "R" + (rows_index + 1).ToString() + "C" + (columns_index + 1).ToString();
                }
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            
            string expression = dataGridView1.CurrentCell.Value.ToString();
            var pattern = @"[R]\d[C]\d";
            var stringVariableMatches = Regex.Replace(expression, pattern,
                m => Cell_Dictionary[m.Value][1]);
            Expression express_ion = new Expression(stringVariableMatches);
            Expression express_ion1 = new Expression(expression);
            string value_of_cell = express_ion.calculate().ToString();
            Cell_Dictionary["R" + (e.RowIndex + 1).ToString() + "C" + (e.ColumnIndex + 1).ToString()][1] = value_of_cell;
            dataGridView1.CurrentCell.Value = Cell_Dictionary["R" + (e.RowIndex + 1).ToString() + "C" + (e.ColumnIndex + 1).ToString()][1];

            string text_expression = express_ion1.getExpressionString().ToString();
            Cell_Dictionary["R" + (e.RowIndex + 1).ToString() + "C" + (e.ColumnIndex + 1).ToString()][0] = text_expression;
        }
    }
}