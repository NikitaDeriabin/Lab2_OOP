using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel
{
   public partial class MyCell : DataGridViewTextBoxCell
    {
        public string NameCell { get; private set; }// имя клетки
        public int ColNameCell { get; private set; }
        public int RowNameCell { get; private set; }
        public string val;// кінцеве значення
        public string exp;// вираз
        public List<MyCell> depends = new List<MyCell>();
        public List<MyCell> dependsOnMe = new List<MyCell>();
        public List<MyCell> new_depends = new List<MyCell>();


        public MyCell(string name, int row, int col)
        {
            NameCell = name;
            RowNameCell = row;
            ColNameCell = col;
            val = "0";
            exp = "";
        }

        public new string Value
        {
            get { return val; }
            set { val = value; }
        }

        public string Exp
        {
            get { return exp; }
            set { exp = value; }
        }

        public void SetCell(string value, string expression, List<MyCell> references, List<MyCell> pointers)// references - мои зависимости; pointers - зависимые от меня
        {
            this.Value = value;
            this.Exp = expression;

            this.depends.Clear();
            this.depends.AddRange(references);

            this.dependsOnMe.Clear();
            this.dependsOnMe.AddRange(pointers);
        }

        public bool CheckForCycle(List<MyCell> check_list) //check_list - depends
        {
            foreach (MyCell check in check_list)
                if (check.NameCell == this.NameCell)
                    return false;
            foreach (MyCell checkOnMe in dependsOnMe)
            {
                foreach (MyCell check in check_list)
                    if (check.NameCell == checkOnMe.NameCell)
                        return false;
                if (!checkOnMe.CheckForCycle(check_list))
                    return false;
            }
            return true;
                
        }

        public void AddDependsOnMeAndDepends()
        {
            foreach (MyCell depend in new_depends)
            {
                depend.dependsOnMe.Add(this);
            }
            depends = new_depends;
        }

        public void DellDependsOnMeAndDepends()
        {
            if(depends != null)
            {
                foreach (MyCell depend in depends)
                    depend.dependsOnMe.Remove(this);
            }
            depends = null;
        }

    }
}
