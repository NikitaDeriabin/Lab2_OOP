using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace Excel
{
    public partial class Excel : Form
    {
        const int sz = 200;
        const int startRow = 10;
        const int startCol = 10;
        int currRow, currCol;
        Grid Gr = new Grid();
        CreateExtraEl ExtraEl = new CreateExtraEl();
       
        Parser pars = new Parser();
        public Excel()
        {
            InitializeComponent();
            CreateTable(startCol, startRow);
           
        }
         
        private void CreateTable(int startCol, int startRow)
        {
            string nameCol = "";

            for (int i = 0; i < startCol; i++)
            {
                nameCol += Grid.to26Sys(i);
                dataGridView1.Columns.Add(nameCol, nameCol);
                nameCol = "";
            }

            dataGridView1.RowCount = startRow;

            for(int i = 0; i < startRow; i++)
            {
                dataGridView1.Rows[i].HeaderCell.Value = i.ToString();
            }

            SetMyCells(startCol, startRow);
        }

        private void SetMyCells(int startCol, int startRow)
        {
            Gr.RowCount = startRow;
            Gr.ColCount = startCol;

            for(int i = 0; i < sz; i++)
            {

                for (int j = 0; j < sz; j++)
                {
                    string name = Grid.to26Sys(j) + i.ToString();
                    Gr.TableCells[i, j] = new MyCell(name, i, j);
                }              
            }

            for(int i = 0; i < startRow; i++)
            {
                for(int j = 0; j < startCol; j++)
                {
                    string nameCell = Grid.to26Sys(j) + i.ToString();
                    Gr.dictionary.Add(nameCell, "");
                }
                
            }
        }

       
        private void Add_Row_Click(object sender, EventArgs e)
        {
            DataGridViewRow row = new DataGridViewRow();
            if(dataGridView1.Columns.Count == 0)
            {
                MessageBox.Show("There are no columns.");
                return;
            }

            dataGridView1.Rows.Add(row);

            dataGridView1.Rows[dataGridView1.Rows.Count - 1].HeaderCell.Value = (dataGridView1.Rows.Count - 1).ToString();

            Gr.AddRow(dataGridView1);
        }

        private void Add_Column_Click(object sender, EventArgs e)
        {
            string colName = Grid.to26Sys(Gr.ColCount);
            dataGridView1.Columns.Add(colName, colName);

            Gr.AddColumn(dataGridView1);
        }

        private void Delete_Row_Click(object sender, EventArgs e)
        {
            int curRow = Gr.RowCount - 1;
            if (!Gr.DeleteRow(dataGridView1))
                return;
            dataGridView1.Rows.RemoveAt(curRow);
        }

        private void Delete_Column_Click(object sender, EventArgs e)
        {
            int curCol = Gr.ColCount - 1;
            if (!Gr.DeleteColumn(dataGridView1))
                return;
            dataGridView1.Columns.RemoveAt(curCol);
        }

        private void DataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            string formula = "";
            currRow = dataGridView1.CurrentCell.RowIndex;
            currCol = dataGridView1.CurrentCell.ColumnIndex;
            Gr.TableCells[currRow, currCol].Exp = (string)dataGridView1[currCol, currRow].Value;
            formula = (string)dataGridView1[currCol, currRow].Value;
            if (formula == null) formula = "";
            textBox1.Text = formula;
            if (formula == null) return;
            Gr.ChangeCell(currRow, currCol, formula, dataGridView1);
            
        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            int currRow = dataGridView1.CurrentCell.RowIndex;
            int currCol = dataGridView1.CurrentCell.ColumnIndex;

            if (Gr.TableCells[currRow, currCol].Exp != "")
            {
                textBox1.Text = Gr.TableCells[currRow, currCol].Exp;
                dataGridView1[currCol, currRow].Value = Gr.TableCells[currRow, currCol].Exp; 
            }
            else textBox1.Text = "";
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            currRow = dataGridView1.CurrentCell.RowIndex;
            currCol = dataGridView1.CurrentCell.ColumnIndex;
            textBox1.Text = Gr.TableCells[currRow, currCol].Exp;
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                textBox1_Leave(sender, e);
        }

        private void SaveFile_Click(object sender, EventArgs e)
        {
           SaveFile();
        }

        private void SaveFile()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Text files(*.txt)|*.txt|All files(*.*)|*.*";
            saveFileDialog.Title = "Save table";
            saveFileDialog.ShowDialog();

            if (saveFileDialog.FileName != "")
            {
                FileStream fs = (FileStream)saveFileDialog.OpenFile();

                StreamWriter sw = new StreamWriter(fs);

                Gr.Save(sw);

                sw.Close();
                fs.Close();
            }
        }

        private void OpenFile_Click(object sender, EventArgs e)
        {
            OpenFile();
        }

        private void OpenFile()
        {
            DialogResult resMessage =  MessageBox.Show("Do you want to save this file?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if(resMessage == DialogResult.Yes)
            {
                SaveFile();
            }

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Text files(*.txt)|*.txt|All files(*.*)|*.*";
            openFileDialog.Title = "Select File";
            if (openFileDialog.ShowDialog() != DialogResult.OK)
                return;
            StreamReader sr = new StreamReader(openFileDialog.FileName);

            try
            {
                Gr.Clear();
                dataGridView1.Rows.Clear();
                dataGridView1.Columns.Clear();

                int row = Convert.ToInt32(sr.ReadLine());
                int col = Convert.ToInt32(sr.ReadLine());

                CreateTable(row, col);
                Gr.Open(row, col, sr, dataGridView1);
            }
            catch(Exception)
            {
                MessageBox.Show("Can't open this file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            sr.Close();
            textBox1.Text = Gr.TableCells[0, 0].Exp;
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            About();
        }

        private void About()
        {
            string about = "Лабораторна робота №2" + '\n' + "Аналог Excel" + '\n' + "Виконав студент групи К-24 : Дерябін Микита";
            about += '\n' + "Мої варіанти : 1(+, -, /, *), 4(^), 5(inc, dec), 6(max(x,y), min(x,y))";
            MessageBox.Show(about, "About", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Excel_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult res = MessageBox.Show("Do you want to save file?", "Warning", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
            if (res == DialogResult.Yes)
                SaveFile();
            if (res == DialogResult.Cancel) e.Cancel = true;
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            currRow = dataGridView1.CurrentCell.RowIndex;
            currCol = dataGridView1.CurrentCell.ColumnIndex;
            string formula = textBox1.Text;

            Gr.ChangeCell(currRow, currCol, formula, dataGridView1);

        }

    }

    
}
