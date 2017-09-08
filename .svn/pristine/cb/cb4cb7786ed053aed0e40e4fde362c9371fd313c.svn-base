using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FlowDemo_V1
{
    public static class Extension
    {
        public static DataGridViewRow CloneWithValues(this DataGridViewRow row)
        {
            DataGridViewRow clonedRow = (DataGridViewRow)row.Clone();
            for (Int32 index = 0; index < row.Cells.Count; index++)
            {
                clonedRow.Cells[index].Value = row.Cells[index].Value;
            }
            return clonedRow;
        }

        public static List<DataGridViewRow> Sort(this DataGridViewSelectedRowCollection rows)
        {
            List<DataGridViewRow> sortedRows = new List<DataGridViewRow>();

            for (int i = 0; i < rows.Count; i++)
            {
                for (int j = sortedRows.Count - 1; j >= 0; j--)
                {
                    if (rows[i].Index > sortedRows[j].Index)
                    {
                        sortedRows.Insert(j + 1, rows[i]);
                        break;
                    }

                    if (j == 0)
                    {
                        sortedRows.Insert(j, rows[i]);
                    }
                }

                if (i == 0)
                {
                    sortedRows.Insert(i, rows[i]);
                }
            }

            return sortedRows;
        }
    }
}
