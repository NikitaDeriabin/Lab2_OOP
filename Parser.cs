using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Excel
{
    class Parser
    {

        public static void Replace(ref string s)
        {
            Regex regexMax = new Regex(@"max");
            Regex regexMin = new Regex(@"min");
            Regex regexInc = new Regex(@"inc\(");
            Regex regexDec = new Regex(@"dec\(");
            Regex regexSpace = new Regex(@"\s");
            string targetMax = "&";
            string targetMin = "#";
            string targetInc = "(1+0";
            string targetDec = "(-1+0";
            string targetEmp = "";

            Regex regexMx = new Regex(@"&");
            Regex regexMn = new Regex(@"#");

            Replacement(ref s, regexSpace, targetEmp);
            Replacement(ref s, regexMax, targetMax);
            Replacement(ref s, regexMin, targetMin);
            Replacement(ref s, regexInc, targetInc);
            Replacement(ref s, regexDec, targetDec);

            InsertMaxMin(ref s);

            Replacement(ref s, regexMx, targetEmp);
            Replacement(ref s, regexMn, targetEmp);

        }

        private static void InsertMaxMin(ref string s)
        {
            Stack<string> st = new Stack<string>();
            StringBuilder sb = new StringBuilder(s);

            for (int i = 0; i < sb.Length; i++)
            {
                if (sb[i] == '&') st.Push("max");
                if (sb[i] == '#') st.Push("min");
                if (sb[i] == ';')
                {
                    if (st.Peek() == "max")
                    {
                        sb[i] = '}';//max
                        st.Pop();
                    }
                    else
                    {
                        sb[i] = '{';//min
                        st.Pop();
                    }
                }
            }
            s = sb.ToString();

        }


        private static void Replacement(ref string s, Regex regex, string target)
        {
            s = regex.Replace(s, target);
        }

        public static double Calc(string s)
        {
            Stack<double> Operands = new Stack<double>();
            Stack<char> Operations = new Stack<char>();
            object token;
            object prevToken = '0';
            int pos = 0;

            s = '(' + s + ')';


            do
            {
                token = GetToken(s, ref pos);
                if (token is char && prevToken is char &&
                    (char)prevToken == '(' && ((char)token == '+' || (char)token == '-'))
                    Operands.Push(0);

                if (token is double)
                {
                    Operands.Push((double)token);
                }

                else if (token is char)
                {
                    if ((char)token == ')')
                    {
                        while (Operations.Count > 0 && Operations.Peek() != '(')
                        {
                            PopFunction(Operands, Operations);
                        }
                        Operations.Pop();
                    }
                    else
                    {
                        while (CanPop((char)token, Operations))
                            PopFunction(Operands, Operations);

                        Operations.Push((char)token);
                    }
                }
                prevToken = token;
            }
            while (token != null);

            if (Operands.Count > 1 || Operations.Count > 0)
                throw new Exception("invalid expression");
            return Operands.Pop();


        }

        private static bool CanPop(char op1, Stack<char> Operations)
        {
            if (Operations.Count == 0)
                return false;
            int p1 = GetPriory(op1);
            int p2 = GetPriory(Operations.Peek());

            return p1 >= 0 && p2 >= 0 && p1 >= p2;
        }

        private static int GetPriory(char op)
        {
            switch (op)
            {
                case '(':
                    return -1;
                case '}':
                case '{':
                    return 4;
                case '+':
                case '-':
                    return 3;
                case '*':
                case '/':
                    return 2;
                case '^':
                    return 1;
                default:
                    throw new Exception("Invalid expression");
            }


        }
        private static void PopFunction(Stack<double> Operands, Stack<char> Operations)
        {
            double A, B;
            if (Operands.Count > 0) B = Operands.Pop();
            else throw new Exception("Invalid expression");
            if (Operands.Count > 0) A = Operands.Pop();
            else throw new Exception("Invalid expression");

            switch (Operations.Pop())
            {
                case '+':
                    Operands.Push(A + B);
                    break;
                case '-':
                    Operands.Push(A - B);
                    break;
                case '*':
                    Operands.Push(A * B);
                    break;
                case '/':
                    if (B != 0)
                        Operands.Push(A / B);
                    else throw new DivideByZeroException("Divide by zero");
                    break;
                case '^':
                    Operands.Push(Math.Pow(A, B));
                    break;
                case '{':
                    Operands.Push(Math.Min(A, B));
                    break;
                case '}':
                    Operands.Push(Math.Max(A, B));
                    break;
            }
        }

        private static object GetToken(string s, ref int pos)
        {
            if (pos == s.Length)
                return null;
            if (char.IsDigit(s[pos]))
                return Convert.ToDouble(ReadDouble(s, ref pos));
            else
                return ReadFunction(s, ref pos);

        }

        private static object ReadDouble(string s, ref int pos)
        {
            string res = "";
            while (pos < s.Length && (char.IsDigit(s[pos]) || s[pos] == ','))
            {
                res += s[pos++];
            }
            return res;
        }

        private static object ReadFunction(string s, ref int pos)
        {
            return s[pos++];
        }

    }

}
