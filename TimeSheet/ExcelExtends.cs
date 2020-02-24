using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NetOffice.ExcelApi
{
    public static class ExcelExtends
    {
        public static string GetString(this Range cell)
        {
            return cell.Value as string;
        }

        public static int GetInt(this Range cell)
        {
            if(cell.Value is decimal v1)
            {
                return decimal.ToInt32(v1);
            }
            else if (cell.Value is double v2)
            {
                return (int)v2;
            }
            else if (cell.Value is int v3)
            {
                return v3;
            }
            return 0;
        }

        public static double? GetDouble(this Range cell)
        {
            if (cell.Value is double v)
            {
                return v;
            }
            return null;
        }

        public static DateTime? GetDateTime(this Range cell)
        {
            if(cell.Value is DateTime d)
            {
                return d;
            }
            return null;
        }
    }
}
