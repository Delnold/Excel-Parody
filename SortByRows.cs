using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Excel_Parody
{
    class SortByRows : IComparer<string>
    {
        public int Compare(string x, string y)
        {
            var first_str = Regex.Matches(x, @"\d+");
            var second_str = Regex.Matches(y, @"\d+");  
            var f_num = Convert.ToDouble(first_str[0].Value) + (Convert.ToDouble(first_str[1].Value) / (1000 ^ (first_str[1].Value.ToString().Length)));
            var s_num = Convert.ToDouble(second_str[0].Value) + (Convert.ToDouble(second_str[1].Value) / (1000 ^ (second_str[1].Value.ToString().Length)));
            if (f_num == s_num)
            {
                return 0;
            }
            if (f_num < s_num)
            {
                return -1;
            }
            else
            {
                return 1;
            }

        }
    }

}
