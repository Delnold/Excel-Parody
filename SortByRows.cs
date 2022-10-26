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
            var f_num = Convert.ToInt32(first_str[0].ToString() + first_str[1].ToString());
            var s_num = Convert.ToInt32(second_str[0].ToString() + second_str[1].ToString());
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
