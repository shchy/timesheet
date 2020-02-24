using NetOffice.ExcelApi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TimeSheet
{
    class Order
    {
        public string OrderID { get; set; }
        public string AssignID { get; set; }
        public string AssignName { get; set; }
        public string Company { get; set; }
        public string Name { get; set; }
        public DateTime Date { get; set; }
        public int Price { get; set; }
        public TimeSpan UnderTime { get; set; }
        public TimeSpan UpperTime { get; set; }
        public int UnderPriceByHour { get; set; }
        public int UpperPriceByHour { get; set; }
        public double WorkPerMonth { get; set; }
    }

    class OrderFile
    {
        const int START_ROW = 6;
        const int ORDER_NO_COL = 3;
        const int ASSIGN_ID_COL = 20;
        const int ASSIGN_NAME_COL = 21;
        const int COMPANY_COL = 23;
        const int NAME_COL = 24;
        const int DATE_COL = 26;
        const int PRICE_COL = 31;
        const int UNDER_TIME_COL = 32;
        const int UPPER_TIME_COL = 33;
        const int UNDER_PRICE_UNIT_COL = 34;
        const int UPPER_PRICE_UNIT_COL = 35;
        const int PER_MONTH_COL = 38;

        public IEnumerable<Order> Read(Worksheet sheet)
        {
            sheet.Activate();
            foreach (var row in sheet.Rows.Skip(START_ROW))
            {
                var orderID = row.Cells[ORDER_NO_COL].GetString();
                if (string.IsNullOrWhiteSpace(orderID))
                {
                    break;
                }
                var assignID = row.Cells[ASSIGN_ID_COL].GetString();
                var assignName = row.Cells[ASSIGN_NAME_COL].GetString();
                var company = row.Cells[COMPANY_COL].GetString();
                var name = row.Cells[NAME_COL].GetString();
                var date = row.Cells[DATE_COL].GetDateTime();
                var price = row.Cells[PRICE_COL].GetInt();
                var underTime = TimeSpan.FromHours(row.Cells[UNDER_TIME_COL].GetDouble().Value);
                var upperTime = TimeSpan.FromHours(row.Cells[UPPER_TIME_COL].GetDouble().Value);
                var underPriceUnit = row.Cells[UNDER_PRICE_UNIT_COL].GetInt();
                var upperPriceUnit = row.Cells[UPPER_PRICE_UNIT_COL].GetInt();
                var perMonth = row.Cells[PER_MONTH_COL].GetDouble();

                yield return new Order
                {
                    OrderID = orderID,
                    AssignID = assignID,
                    AssignName = assignName,
                    Company = company,
                    Name = name,
                    Date = date.Value,
                    Price = price,
                    UnderTime = underTime,
                    UpperTime = upperTime,
                    UnderPriceByHour = underPriceUnit,
                    UpperPriceByHour = upperPriceUnit,
                    WorkPerMonth = perMonth.Value,
                };
            }
        }
    }

    
}
