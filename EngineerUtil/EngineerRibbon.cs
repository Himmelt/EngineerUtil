using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace EngineerUtil
{
    public partial class EngineerRibbon
    {

        private Application application;

        private void EngineerRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            application = Globals.ThisAddIn.Application;
        }

        private void btnPLC_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook workbook = application.ActiveWorkbook;
            try
            {
                Worksheet plcSheet = workbook.Sheets["PLC_Engineers_Plan"];
                Worksheet plcSheet2 = workbook.Sheets["PLC_Engineers_Resume"];
                Range Cells = plcSheet.Cells;
                int startRow = 0, startCol = 0, endRow = 0, endCol = 0;

                if (plcSheet != null && plcSheet2 != null)
                {
                    bool findRange = false;
                    for (int i = 1; i <= 800; i++)
                    {
                        for (int j = 1; j <= 100; j++)
                        {
                            //Cells[i, j].Value2 = Convert.ToString(i + j);
                            var value = Cells[i, j].Value2;
                            if (value!=null && value.ToString().Equals("MARK_START"))
                            {
                                startRow = i;
                                startCol = j;
                            }
                            else if (value != null && value.ToString().Equals("MARK_END"))
                            {
                                endRow = i;
                                endCol = j;
                                findRange = true;
                                break;
                            }
                        }
                        if (findRange) break;
                    }

                    if (findRange)
                    {
                        Range Cells2 = plcSheet2.Cells;
                        int k = 1;
                        var name = Cells2[15 * (k - 1) + 1, 3].Value2;
                        
                        while(name !=null )
                        {
                            string theName = name.ToString();

                            Dictionary<int, List<int>> map = new Dictionary<int, List<int>>();

                            for(int m = startRow; m <= endRow; m++)
                            {
                                for(int n = startCol; n <= endCol; n++)
                                {
                                    var value = Cells[m, n].Value2;
                                    if (value != null && value.ToString().Equals(theName))
                                    {
                                        List<int> list = null;
                                        if (map.TryGetValue(m, out list)) {
                                            map[m].Add(n);
                                        }
                                        else
                                        {
                                            list = new List<int>();
                                            list.Add(n);
                                            map.Add(m,list);
                                        }
                                    }
                                }
                            }


                            int l = 1;
                            // 先把14行清空
                            for (l = 1; l <= 14; l++)
                            {
                                Cells2[15 * (k - 1) + 1 + l, 4].Value2 = null;
                                Cells2[15 * (k - 1) + 1 + l, 5].Value2 = null;
                            }

                            l = 1;

                            var sort = from obj in map orderby obj.Key ascending select obj;

                            foreach (KeyValuePair<int, List<int>> kvp in sort)
                            {
                                if (l <= 14)
                                {
                                    int line = kvp.Key;
                                    List<int> list = kvp.Value;
                                    string result = "KW " + list.First() + "-" + list.Last();
                                    Cells2[15 * (k - 1) + 1 + l, 4].Value2 = result;
                                    Range range = Cells[line, startCol + 3];
                                    Cells2[15 * (k - 1) + 1 + l, 5].Value2 = range.MergeArea.Cells[1, 1].Value2;
                                }
                                l++;
                            }

                            k++;
                            name = Cells2[15 * (k - 1) + 1, 3].Value2;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Not found valid range mark !!");
                    }
                }
                else
                {
                    MessageBox.Show("PLC Engineers workSheet not exist !!");
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("PLC Engineers workSheet not exist !!");
                MessageBox.Show($"Error message: {ex.Message}\n" + $"Details:\n{ex.StackTrace}");
            }
        }
    }
}
