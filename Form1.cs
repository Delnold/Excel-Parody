using System.Data;
using System.Text.RegularExpressions;
using System.Linq;
using System.Collections.Generic;
using System.Collections;
using System.Windows.Forms;
using org.mariuszgromada.math.mxparser;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Drawing;
using System.Text.Json.Serialization;
using Newtonsoft.Json;



namespace Excel_Parody
{

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private static DataTable table = new DataTable();
        private static SortedDictionary<string, string[]> Cell_Dictionary = new SortedDictionary<string, string[]>(new SortByRows());
        public static void convert_From_Dict_To_Table()
        {
            table.Rows.Clear();
            table.Columns.Clear();
                var index_of_R_C = Cell_Dictionary.Keys.Last();
                var parsing_Row_Col = Regex.Matches(index_of_R_C, @"\d+");
                var rows_c = Int32.Parse(parsing_Row_Col[0].ToString());
                var col_c = Int32.Parse(parsing_Row_Col[1].ToString());

                for (int i_c = 1; i_c <= col_c; i_c++)
                {
                    table.Columns.Add((table.Columns.Count + 1).ToString());
                }
                for (int i_r = 1; i_r <= rows_c; i_r++)
                {
                    table.Rows.Add();
                }
           
            convert_From_Expression_To_Value();

        }
        public static void convert_From_Expression_To_Value()
        {
            var index_of_R_C = Cell_Dictionary.Keys.Last();
            var parsing_Row_Col = Regex.Matches(index_of_R_C, @"\d+");
            var rows_c = Int32.Parse(parsing_Row_Col[0].ToString());
            var col_c = Int32.Parse(parsing_Row_Col[1].ToString());
            dynamic expression_RegEX = "";
            dynamic mx_type_Expression;
            string expression;
            string pattern = @"[R]\d+[C]\d+";
            bool success = false;

            for (int i_r = 1; i_r <= rows_c; i_r++)
            {
                for (int i_c = 1; i_c <= col_c; i_c++)
                {
                    if (Cell_Dictionary["R" + (i_r).ToString() + "C" + (i_c).ToString()][1] != "")
                    {
                        expression = Cell_Dictionary["R" + (i_r).ToString() + "C" + (i_c).ToString()][1];

                        try
                        {
                            if (Regex.Matches(expression, @"R" + (i_r).ToString() + @"C" + (i_c).ToString()).Count > 0)
                            {
                                Cell_Dictionary["R" + (i_r).ToString() + "C" + (i_c).ToString()][0] = "RecursionRef";
                                table.Rows[i_r - 1][i_c - 1] = Cell_Dictionary["R" + (i_r).ToString() + "C" + (i_c).ToString()][0];
                                success = false;
                                continue;
                            }

                            expression_RegEX = Regex.Replace(expression, pattern,
                            m => Cell_Dictionary[m.Value][0]);

                            if (expression_RegEX.Contains("RecursionRef"))
                            {
                                Cell_Dictionary["R" + (i_r).ToString() + "C" + (i_c).ToString()][0] = "Rec_Referecnce";
                                table.Rows[i_r - 1][i_c - 1] = Cell_Dictionary["R" + (i_r).ToString() + "C" + (i_c).ToString()][0];
                                success = false;
                                continue;
                            }
                            success = true;
                            if (expression_RegEX.Contains("Inv_Ind_Format"))
                            {
                                Cell_Dictionary["R" + (i_r).ToString() + "C" + (i_c).ToString()][0] = "Inv_Ind_Format_Ref";
                                table.Rows[i_r - 1][i_c - 1] = Cell_Dictionary["R" + (i_r).ToString() + "C" + (i_c).ToString()][0];
                                success = false;
                                continue;
                            }
                            success = true;

                        }
                        catch (KeyNotFoundException)
                        {
                             Cell_Dictionary["R" + (i_r).ToString() + "C" + (i_c).ToString()][0] = "Inv_Ind_Format";
                             table.Rows[i_r - 1][i_c - 1] = Cell_Dictionary["R" + (i_r).ToString() + "C" + (i_c).ToString()][0];
               

                        }
                        finally
                        {
                            if (success)
                            {
                                
                                mx_type_Expression = new Expression(expression_RegEX);
                                if (Double.IsNaN(mx_type_Expression.calculate()))
                                {
                                    Cell_Dictionary["R" + (i_r).ToString() + "C" + (i_c).ToString()][0] = "NaN";
                                    table.Rows[i_r - 1][i_c - 1] = Cell_Dictionary["R" + (i_r).ToString() + "C" + (i_c).ToString()][0];
                                    
                                }
                                else
                                {
                                    Cell_Dictionary["R" + (i_r).ToString() + "C" + (i_c).ToString()][0] = mx_type_Expression.calculate().ToString();
                                    table.Rows[i_r - 1][i_c - 1] = Cell_Dictionary["R" + (i_r).ToString() + "C" + (i_c).ToString()][0];
                                }
                            }
                        }
                       
                    }
                    else
                        table.Rows[i_r - 1][i_c - 1] = Cell_Dictionary["R" + (i_r).ToString() + "C" + (i_c).ToString()][0];
                }
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            for (int i = 1; i <= 7; i++)
            {
                for (int a = 1; a <= 7; a++)
                {
                    Cell_Dictionary.Add("R" + i.ToString() + "C" + a.ToString(), new[] { "", "" });
                }
            }
            convert_From_Dict_To_Table();

            dataGridView1.DataSource = table;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        // Add Row
        private void button1_Click(object sender, EventArgs e)
        {

            var index_of_R_C = Cell_Dictionary.Keys.Last();
            var parsing_Row_Col = Regex.Matches(index_of_R_C, @"\d+");
            var rows_c = Int32.Parse(parsing_Row_Col[0].ToString());
            var col_c = Int32.Parse(parsing_Row_Col[1].ToString());

            for (int i = 1; i <= col_c; i++)
            {
                Cell_Dictionary.Add("R" + (rows_c + 1).ToString() + "C" + i.ToString(), new[] { "", "" });
            }
            convert_From_Dict_To_Table();
        }
        // Add Column
        private void button2_Click(object sender, EventArgs e)
        {
            
            var index_of_R_C = Cell_Dictionary.Keys.Last();
            var parsing_Row_Col = Regex.Matches(index_of_R_C, @"\d+");
            var rows_c = Int32.Parse(parsing_Row_Col[0].ToString());
            var col_c = Int32.Parse(parsing_Row_Col[1].ToString());

            for (int i = 1; i <= rows_c; i++)
            {
                Cell_Dictionary.Add("R" + (i).ToString() + "C" + (col_c + 1).ToString(), new[] { "", "" });
            }
            convert_From_Dict_To_Table();
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
            
            var index_of_R_C = Cell_Dictionary.Keys.Last();
            var parsing_Row_Col = Regex.Matches(index_of_R_C, @"\d+");
            var rows_c = Int32.Parse(parsing_Row_Col[0].ToString());
            var col_c = Int32.Parse(parsing_Row_Col[1].ToString());
            if(col_c == 1)
            {
                MessageBox.Show("You can`t delete last column!");
                return;
            } 
                
            for (int a = 1; a <= rows_c; a++)
                {
                    for (int i = column_ind + 1; i < col_c; i++)
                    {
                        Cell_Dictionary["R" + (a.ToString()) + "C" + (i.ToString())][0] = Cell_Dictionary["R" + (a.ToString()) + "C" + (i + 1).ToString()][0];
                        Cell_Dictionary["R" + (a.ToString()) + "C" + (i.ToString())][1] = Cell_Dictionary["R" + (a.ToString()) + "C" + (i + 1).ToString()][1];
                    }

                }
                for (int i = 1; i <= rows_c; i++)
                {
                    Cell_Dictionary.Remove("R" + (i).ToString() + "C" + (col_c).ToString());
                }
            
            convert_From_Dict_To_Table();


        }

        // Delete Row
        private void button1_Click_1(object sender, EventArgs e)
        {
            int rows_ind = dataGridView1.CurrentCell.RowIndex;
            
            var index_of_R_C = Cell_Dictionary.Keys.Last();
            var parsing_Row_Col = Regex.Matches(index_of_R_C, @"\d+");
            var rows_c = Int32.Parse(parsing_Row_Col[0].ToString());
            var col_c = Int32.Parse(parsing_Row_Col[1].ToString());
            if (rows_c == 1)
            {
                MessageBox.Show("You can`t delete last row!");
                return;
            }
            for (int a = 1; a <= col_c; a++)
            {
                for (int i = rows_ind + 1; i < rows_c; i++)
                {
                    Cell_Dictionary["R" + (i.ToString()) + "C" + (a.ToString())][0] = Cell_Dictionary["R" + (i + 1).ToString() + "C" + (a.ToString())][0];
                    Cell_Dictionary["R" + (i.ToString()) + "C" + (a.ToString())][1] = Cell_Dictionary["R" + (i + 1).ToString() + "C" + (a.ToString())][1];
                }

            }
            for (int i = 1; i <= col_c; i++)
            {
                Cell_Dictionary.Remove("R" + (rows_c).ToString() + "C" + (i.ToString()));
            }
            convert_From_Dict_To_Table();

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
            
            int rows_ind = dataGridView1.CurrentCell.RowIndex;
            int cols_ind = dataGridView1.CurrentCell.ColumnIndex;
            textBox1.Text = Cell_Dictionary["R" + (rows_ind + 1).ToString() + "C" + (cols_ind + 1).ToString()][1];
            textBox2.Text = "R" + (rows_ind + 1).ToString() + "C" + (cols_ind + 1).ToString();
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
          

            string expression = dataGridView1.CurrentCell.Value.ToString();
            /*
            string pattern = @"[R]\d+[C]\d+";
            dynamic expression_RegEX;
            dynamic mx_type_Expression;
            */
            int rows_ind = dataGridView1.CurrentCell.RowIndex;
            int cols_ind = dataGridView1.CurrentCell.ColumnIndex;
            Cell_Dictionary["R" + (rows_ind + 1).ToString() + "C" + (cols_ind + 1).ToString()][1] = expression;
            convert_From_Dict_To_Table();
            /*
            try
            {
                expression_RegEX = Regex.Replace(expression, pattern,
                        m => Cell_Dictionary[m.Value][0]);
            }
            catch
            {
                Cell_Dictionary["R" + (rows_ind + 1).ToString() + "C" + (cols_ind + 1).ToString()][0] = "InvalidIndexFormat";
                Cell_Dictionary["R" + (rows_ind + 1).ToString() + "C" + (cols_ind + 1).ToString()][1] = expression;
                convert_From_Dict_To_Table();
                return;
            }
            mx_type_Expression = new Expression(expression_RegEX);

            Cell_Dictionary["R" + (rows_ind + 1).ToString() + "C" + (cols_ind + 1).ToString()][0] = mx_type_Expression.calculate().ToString();
            Cell_Dictionary["R" + (rows_ind + 1).ToString() + "C" + (cols_ind + 1).ToString()][1] = expression;

            convert_From_Dict_To_Table();
            */
        }
        // Save
        private void button3_Click(object sender, EventArgs e)
        {
            File.WriteAllText(@"D:\save_test.json", JsonConvert.SerializeObject(Cell_Dictionary));
        }
        // Open
        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "json files (*json)|*.json|All files (*.*)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                Cell_Dictionary = JsonConvert.DeserializeObject<SortedDictionary<string, string[]>>(File.ReadAllText(ofd.FileName));
                convert_From_Dict_To_Table();
            }
;        }
    }
}