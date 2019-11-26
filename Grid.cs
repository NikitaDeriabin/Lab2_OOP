using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.IO;


namespace Excel
{
    class Grid
    {
        private const int sz = 200;
        private const int initColCount = 10;
        private const int initRowCount = 10;
        public int ColCount;
        public int RowCount;
        public Dictionary<string, string> dictionary = new Dictionary<string, string>();
        public MyCell[,] TableCells = new MyCell[sz, sz];

        public void Clear()
        {
           for(int i = 0; i < sz; i++)
           {
                for(int j = 0; j < sz; j++)
                {
                    string name = FullName(i, j);
                    TableCells[i, j] = new MyCell(name, i, j);
                }
           }

            dictionary.Clear();
            RowCount = 0;
            ColCount = 0;
        }
        public static string to26Sys(int i)
        {
            int k = 0;
            int[] Arr = new int[100];
            while (i > 25)
            {
                Arr[k++] = i / 26 - 1;
                i = i % 26;
            }
            Arr[k] = i;
            string res = "";
            for (int j = 0; j <= k; j++)
            {
                res += ((char)('A' + Arr[j])).ToString();
            }

            return res;
        }

        public static int[] from26Sys(string index)
        {
            int[] nums = new int[2];
            StringBuilder first_part = new StringBuilder();
            int letter_index = 0;
            foreach (char c in index)
            {
                if (Char.IsLetter(c))
                {
                    first_part.Append(c);
                    letter_index++;
                    continue;
                }
                string first = first_part.ToString();

                char[] charArray = first.ToCharArray();
                int len = charArray.Length;
                int res = 0;
                for (int i = len - 2; i >= 0; i--)
                {
                    res += (((int)charArray[i] - (int)'A') + 1) * Convert.ToInt32(Math.Pow(26, len - i - 1));
                }

                res += ((int)charArray[len - 1] - (int)'A');
                nums[0] = res;
                break;
            }
            nums[1] = Convert.ToInt32(index.Substring(letter_index));
            return nums;
        }

        public void ChangeCell(int currRow, int currCol, string formula, DataGridView dataGridView1)
        {

            TableCells[currRow, currCol].DellDependsOnMeAndDepends();
            TableCells[currRow, currCol].Exp = formula;

            // set new depends
            TableCells[currRow, currCol].new_depends.Clear();

            string new_formula = ConvertDepends(currRow, currCol, formula);
            formula = new_formula;

            //check for cycle
            if (!TableCells[currRow, currCol].CheckForCycle(TableCells[currRow, currCol].new_depends))
            {
                MessageBox.Show("There is a cycle of cells!");
                TableCells[currRow, currCol].Value = "Error";
                dictionary[FullName(currRow, currCol)]= "Error";
                dataGridView1[currCol, currRow].Value = "Error";


                foreach (MyCell cell in TableCells[currRow, currCol].dependsOnMe)
                {
                    UpdateCells(cell, dataGridView1);
                }

                TableCells[currRow, currCol].AddDependsOnMeAndDepends();

                return;
            }

            //добавляем клетку в зависящие клетки
            TableCells[currRow, currCol].AddDependsOnMeAndDepends();

            GetResult(new_formula, currRow, currCol, dataGridView1);

        }

        public void GetResult(string formula, int currRow, int currCol, DataGridView dataGridView1)
        {
            try
            {
                double result;
                if (formula == null)
                {
                    return;
                }
                Parser.Replace(ref formula);
                if (formula.Length != 0)
                {
                    result = Parser.Calc(formula);
                    TableCells[currRow, currCol].Value = result.ToString();
                    dataGridView1[currCol, currRow].Value = result.ToString();
                    dictionary[FullName(currRow, currCol)] = result.ToString();
                }
                else
                {
                    TableCells[currRow, currCol].Value = formula;
                    dataGridView1[currCol, currRow].Value = formula;
                    dictionary[FullName(currRow, currCol)] = "0";
                }
            }
            catch (DivideByZeroException)
            {
                dataGridView1[currCol, currRow].Value = "Divide by zero";
                TableCells[currRow, currCol].Value = "Error";
                dictionary[FullName(currRow, currCol)] = "Error";
            }
            catch (Exception)
            {
                dataGridView1[currCol, currRow].Value = "Error";
                TableCells[currRow, currCol].Value = "Error";
                dictionary[FullName(currRow, currCol)] = "Error";
            }

            foreach (MyCell cell in TableCells[currRow, currCol].dependsOnMe)
            {
                UpdateCells(cell, dataGridView1);
            }
        }

        public bool UpdateCells(MyCell cell, DataGridView dataGridView1)
        {
            cell.new_depends.Clear();

            string new_exp = ConvertDepends(cell.RowNameCell, cell.ColNameCell, cell.Exp);
            GetResult(new_exp, cell.RowNameCell, cell.ColNameCell, dataGridView1) ;

            return true;

        }

        public string FullName(int row, int col)
        {
            return to26Sys(col) + row;
        }

        public string ConvertDepends(int row, int col, string exp)
        {
            Regex regex = new Regex(@"[A-Z]+[0-9]+");
            int[] nums;

            foreach(Match match in regex.Matches(exp))
            {
                if(dictionary.ContainsKey(match.Value))
                {
                    nums = from26Sys(match.Value);
                    TableCells[row, col].new_depends.Add(TableCells[nums[1], nums[0]]);
                }
            }

            MatchEvaluator myEvaluator = new MatchEvaluator(DepToValue);
            string new_exp = regex.Replace(exp, myEvaluator);
            return new_exp;
        }

        public string DepToValue(Match m)
        {
            if(dictionary.ContainsKey(m.Value))
            {
                if (dictionary[m.Value] == "")
                    return "0";
                else return dictionary[m.Value];

            }
            return m.Value;
        }

        public void UpdateDepends()
        {
            foreach (MyCell cell in TableCells)
            {
                if (cell.depends != null)
                    cell.depends.Clear();
                if (cell.new_depends != null)
                    cell.new_depends.Clear();

                if (cell.Exp == "")
                    continue;

                string new_exp = cell.Exp;
                new_exp = ConvertDepends(cell.RowNameCell, cell.ColNameCell, cell.Exp);
                cell.depends.AddRange(cell.new_depends);
            }
        }

        public void AddRow(DataGridView dgv)
        {
            RowCount++;
            for(int i = 0; i < ColCount; i++)
            {
                string name = FullName(RowCount - 1, i);
                dictionary.Add(name, "");
            }

            UpdateDepends();

            foreach(MyCell cell in TableCells)
            {
                if (cell.depends != null)
                    foreach (MyCell cell_in_dep in cell.depends)
                        if (cell_in_dep.RowNameCell == RowCount - 1)
                            if (!cell_in_dep.dependsOnMe.Contains(cell))
                                cell_in_dep.dependsOnMe.Add(cell);
            }

            for(int i = 0; i < ColCount; i++)
            {
                foreach(MyCell cell in TableCells[RowCount - 1, i].dependsOnMe)
                    UpdateCells(cell, dgv);
            }
        }

        public void AddColumn(DataGridView dgv)
        {
            ColCount++;
            for (int i = 0; i < RowCount; i++)
            {
                string name = FullName(i, ColCount - 1);
                dictionary.Add(name, "");
            }

            UpdateDepends();

            foreach (MyCell cell in TableCells)
            {
                if (cell.depends != null)
                    foreach (MyCell cell_in_dep in cell.depends)
                        if (cell_in_dep.ColNameCell == ColCount - 1)
                            if (!cell_in_dep.dependsOnMe.Contains(cell))
                                cell_in_dep.dependsOnMe.Add(cell);
            }

            for (int i = 0; i < RowCount; i++)
            {
                foreach (MyCell cell in TableCells[i, ColCount - 1].dependsOnMe)
                    UpdateCells(cell, dgv);
            }
        }

        public bool DeleteRow(DataGridView dgv)
        {
            if (RowCount == 1)
            {
                MessageBox.Show("You can't delete row", "Warning!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            List<MyCell> notEmptyCells = new List<MyCell>();
            List<MyCell> dependsCells = new List<MyCell>();

            for(int i = 0; i < ColCount; i++)
            {
                string name = FullName(RowCount - 1, i);
                if (dictionary[name] != "0" && dictionary[name] != "")
                    notEmptyCells.Add(TableCells[RowCount - 1, i]);
                if (TableCells[RowCount - 1, i].dependsOnMe.Count != 0)
                    dependsCells.AddRange(TableCells[RowCount - 1, i].dependsOnMe);
            }

            if(notEmptyCells.Count != 0 || dependsCells.Count != 0)
            {
                string errorMessage = "";
                if(notEmptyCells.Count != 0)
                {
                    errorMessage += "There is not empty cells : ";
                    foreach(MyCell cell in notEmptyCells)
                    {
                        string str = cell.NameCell + "; ";
                        errorMessage += str;
                    }

                    errorMessage += "\n";
                }

                if(dependsCells.Count != 0)
                {
                    errorMessage += "There is cells, thats depends from cells in this row : ";
                    foreach(MyCell cell in dependsCells)
                    {
                        string str = cell.NameCell + "; ";
                        errorMessage += str;
                    }

                    errorMessage += "\n";
                }

                errorMessage += "Are you sure to delete this row? ";
                DialogResult result = MessageBox.Show(errorMessage, "Warning!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.No)
                    return false;
            }

            List<MyCell> new_dependCells = new List<MyCell>();
            new_dependCells.AddRange(dependsCells);
            foreach (MyCell cell in dependsCells)
            {
                for (int i = 0; i < ColCount; i++)
                    if (cell.NameCell == TableCells[RowCount - 1, i].NameCell)
                        new_dependCells.Remove(cell);
            }

            for(int i = 0; i < ColCount; i++)
            {
                string name = FullName(RowCount - 1, i);
                dictionary.Remove(name);
            }

            for(int i = 0; i < ColCount; i++)
            {
                if (dgv[i, RowCount - 1].Value == null)
                    continue;
                TableCells[RowCount - 1, i].DellDependsOnMeAndDepends();
            }

            foreach(MyCell cell in new_dependCells)
            {
                UpdateCells(cell, dgv);
            }

            for(int i = 0; i < ColCount; i++)
            {
                string name = FullName(RowCount - 1, i);
                TableCells[RowCount - 1, i] = new MyCell(name, RowCount - 1, i);
            }

            RowCount--;
            return true;
        }

        public bool DeleteColumn(DataGridView dgv)
        {
            if (ColCount == 1)
            {
                MessageBox.Show("You can't delete column", "Warning!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            List<MyCell> notEmptyCells = new List<MyCell>();
            List<MyCell> dependsCells = new List<MyCell>();

            for (int i = 0; i < RowCount; i++)
            {
                string name = FullName(i, ColCount - 1);
                if (dictionary[name] != "0" && dictionary[name] != "")
                    notEmptyCells.Add(TableCells[i, ColCount - 1]);
                if (TableCells[i, ColCount - 1].dependsOnMe.Count != 0)
                    dependsCells.AddRange(TableCells[i, ColCount - 1].dependsOnMe);
            }

            if (notEmptyCells.Count != 0 || dependsCells.Count != 0)
            {
                string errorMessage = "";
                if (notEmptyCells.Count != 0)
                {
                    errorMessage += "There is not empty cells : ";
                    foreach (MyCell cell in notEmptyCells)
                    {
                        string str = cell.NameCell + "; ";
                        errorMessage += str;
                    }

                    errorMessage += "\n";
                }

                if (dependsCells.Count != 0)
                {
                    errorMessage += "There is cells, thats depends from cells in this row : ";
                    foreach (MyCell cell in dependsCells)
                    {
                        string str = cell.NameCell + "; ";
                        errorMessage += str;
                    }

                    errorMessage += "\n";
                }

                errorMessage += "Are you sure to delete this column? ";
                DialogResult result = MessageBox.Show(errorMessage, "Warning!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.No)
                    return false;
            }

            List<MyCell> new_dependCells = new List<MyCell>();
            new_dependCells.AddRange(dependsCells);
            foreach (MyCell cell in dependsCells)
            {
                for (int i = 0; i < RowCount; i++)
                    if (cell.NameCell == TableCells[i, ColCount - 1].NameCell)
                        new_dependCells.Remove(cell);
            }

            for (int i = 0; i < RowCount; i++)
            {
                string name = FullName(i, ColCount - 1);
                dictionary.Remove(name);
            }

            for (int i = 0; i < RowCount; i++)
            {
                if (dgv[ColCount - 1, i].Value == null)
                    continue;
                TableCells[i, ColCount - 1].DellDependsOnMeAndDepends();
            }

            foreach (MyCell cell in new_dependCells)
            {
                UpdateCells(cell, dgv);
            }

            for (int i = 0; i < RowCount; i++)
            {
                string name = FullName(ColCount - 1, i);
                TableCells[i, ColCount - 1] = new MyCell(name, i, ColCount - 1);
            }

            ColCount--;
            return true;
        }

        public void Save(StreamWriter sw)
        {
            sw.WriteLine(RowCount);
            sw.WriteLine(ColCount);

            foreach (MyCell cell in TableCells)
            {
                sw.WriteLine(cell.NameCell);
                sw.WriteLine(cell.Exp);
                sw.WriteLine(cell.Value);

                if (cell.depends == null)
                    sw.WriteLine(0);
                else
                {
                    sw.WriteLine(cell.depends.Count);
                    foreach (MyCell depCell in cell.depends)
                        sw.WriteLine(depCell.NameCell);
                }

                if(cell.dependsOnMe == null)
                    sw.WriteLine(0);

                else
                {
                    sw.WriteLine(cell.dependsOnMe.Count);
                    foreach (MyCell depOnMe in cell.dependsOnMe)
                        sw.WriteLine(depOnMe.NameCell);
                }
            }
        }

        public void Open(int row, int col, StreamReader sr, DataGridView dgv)
        {
            for(int r = 0; r < sz; r++)
            {
                for(int c = 0; c < sz; c++)
                {
                    string index = sr.ReadLine();
                    string expression = sr.ReadLine();
                    string value = sr.ReadLine();

                    if(c < col && r < row)
                    {
                        if (expression != "")
                            dictionary[index] = value;
                        else
                            dictionary[index] = "";
                    }
                    

                    int depCount = Convert.ToInt32(sr.ReadLine());
                    List<MyCell> newDep = new List<MyCell>();
                    string depend; 
                    for(int i = 0; i < depCount; i++)
                    {
                        depend = sr.ReadLine();
                        newDep.Add(TableCells[from26Sys(depend)[1], from26Sys(depend)[0]]);
                    }

                    int depOnMeCount = Convert.ToInt32(sr.ReadLine());
                    List<MyCell> newDepOnMe = new List<MyCell>();
                    string depOnMe;
                    for(int i = 0; i < depOnMeCount; i++)
                    {
                        depOnMe = sr.ReadLine();
                        newDepOnMe.Add(TableCells[from26Sys(depOnMe)[1], from26Sys(depOnMe)[0]]);
                    }
                    
                    TableCells[r, c].SetCell(value, expression, newDep, newDepOnMe);

                    int icol = TableCells[r, c].ColNameCell;
                    int irow = TableCells[r, c].RowNameCell;

                    if (icol >= row || irow >= col) continue;
                    dgv[icol, irow].Value = dictionary[index];
                }
            }
        }


    }
}
