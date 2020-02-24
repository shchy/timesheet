using NetOffice.ExcelApi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TimeSheet
{
    class Worker
    {
        public string Name { get; set; }
        public string Company { get; set; }
        public string AssignName { get; set; }
        public string OrderID { get; set; }
    }

    class WorkRecord
    {
        public DateTime Date { get; set; }
        public TimeSpan? Start { get; set; }
        public TimeSpan? End { get; set; }
        public TimeSpan? BreakTime { get; set; }
        public string Note { get; set; }



        public bool IsWorking() {
            return Start.HasValue
                && End.HasValue;
        }

        public TimeSpan GetWorkTime()
        {
            if (!this.IsWorking())
            {
                return TimeSpan.Zero;
            }
            return End.Value - Start.Value - (BreakTime ?? TimeSpan.Zero);
        }
    }

    class WorkAssign
    {
        public string ID { get; set; }
        public TimeSpan Value { get; set; }
        public int Expense { get; set; }
    }

    class TimeSheet
    {
        public Worker Worker { get; set; }
        public IEnumerable<WorkRecord> WorkRecords { get; set; }
        public IEnumerable<WorkAssign> WorkAssigns { get; set; }

        public TimeSpan GetTotalWorkTime() {
            var sum = TimeSpan.Zero;
            foreach (var record in WorkRecords)
            {
                sum = sum.Add(record.GetWorkTime());
            }
            return sum;
        }
    }

    class TimeSheetFile
    {
        const int nameY = 5;
        const int nameX = 5;
        const int companyY = 4;
        const int companyX = 5;
        const int reasonY = 7;
        const int reasonX = 5;
        const int quoteNoY = 52;
        const int quoteNoX = 25;

        const int workRecordStartY = 13;
        const int workDateX = 2;
        const int workStartTimeX = 6;
        const int workEndTimeX = 9;
        const int workBreakTimeX = 12;
        const int workNoteX = 19;
        const int workAssignStartY = 50;
        const int workAssignIDX = 2;
        const int workAssignTimeX = 9;
        const int workAssignExpenseX = 13;

        public TimeSheet Read(Worksheet sheet)
        {
            sheet.Activate();

            // 名前
            var userName = sheet.Rows[nameY].Cells[nameX].GetString();
            // 所属
            var company = sheet.Rows[companyY].Cells[companyX].GetString();
            // 件名
            var assignName = sheet.Rows[reasonY].Cells[reasonX].GetString();
            // 見積もり番号
            var orderNo = sheet.Rows[quoteNoY].Cells[quoteNoX].GetString();
            var worker = new Worker
            {
                Name = userName,
                AssignName = assignName,
                Company = company,
                OrderID = orderNo,
            };

            var records = GetRecords(sheet).ToArray();

            var assigns = GetAssigns(sheet).ToArray();

            return new TimeSheet
            {
                Worker = worker,
                WorkRecords = records,
                WorkAssigns = assigns,
            };
        }

        private IEnumerable<WorkAssign> GetAssigns(Worksheet sheet)
        {
            foreach (var row in sheet.Rows.Skip(workAssignStartY-1))
            {
                var assignID = row.Cells[workAssignIDX].GetString();
                var workValue = row.Cells[workAssignTimeX].GetDouble();
                var expense = row.Cells[workAssignExpenseX].GetInt();

                if (string.IsNullOrWhiteSpace(assignID)
                    && !workValue.HasValue)
                {
                    break;
                }

                yield return new WorkAssign
                {
                    ID = assignID,
                    Value = TimeSpan.FromHours(workValue.Value),
                    Expense = expense,
                };
            }
        }

        private IEnumerable<WorkRecord> GetRecords(Worksheet sheet)
        {
            foreach (var row in sheet.Rows.Skip(workRecordStartY-1).Take(32))
            {
                var date = row.Cells[workDateX].GetDateTime();
                if (!date.HasValue)
                {
                    break;
                }
                var startTime = row.Cells[workStartTimeX].GetDouble();
                var endTime = row.Cells[workEndTimeX].GetDouble();
                var breakTime = row.Cells[workBreakTimeX].GetDouble();
                var note = row.Cells[workNoteX].GetString();

                yield return new WorkRecord
                {
                    Date = date.Value,
                    Start = startTime.HasValue ? TimeSpan.FromHours(startTime.Value * 24) : default,
                    End = endTime.HasValue ? TimeSpan.FromHours(endTime.Value * 24) : default,
                    BreakTime = breakTime.HasValue ? TimeSpan.FromHours(breakTime.Value * 24) : default,
                    Note = note,
                };
            }
        }
    }

    
}
