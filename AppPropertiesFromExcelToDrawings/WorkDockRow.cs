using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AppPropertiesFromExcelToDrawings
{
    public class WorkDockRow
    {
        public int? Key { get; private set; }
        public string Id { get; private set; }
        public string KitCode { get; private set; } = "";
        public string KitName { get; private set; } = "";
        public string Date { get; private set; }
        public bool? IsIssue { get; private set; }
        public bool? IsChecked { get; set; }
        public bool? IsValidName { get => KitName.Length < 21; }


        public WorkDockRow(Array row, int rowIndex)
        {
            if (row == null)
                throw new ArgumentNullException("Ошибка!");

            Key = rowIndex;

            Id = row.GetValue(0, 0).ToString();

            KitCode = row.GetValue(0, 1).ToString();

            KitName = row.GetValue(0, 2).ToString();

            Date = row.GetValue(0, 3).ToString().Split().First();

            IsIssue = row.GetValue(0, 4) != null ? true : false;
            IsChecked = row.GetValue(0, row.Length - 1) != null ? true : false;
        }
    }
}
