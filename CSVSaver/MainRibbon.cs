using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CSVSaver
{
    public partial class MainRibbon
    {
        //методы событий
        /// <summary> Создание CSV файла с параметрами листов. </summary>
        private void ImportToCSV_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var sheetData = GetActiveSheetDataArray();

                string path = Globals.ThisAddIn.Application.ActiveWorkbook.Path + "\\CSV\\";
                string name = Globals.ThisAddIn.Application.ActiveWorkbook.Name.Split('.').First() + "_" + GetInventory(sheetData) + ".csv";

                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);

                File.WriteAllText(path + name, BuildingContent(sheetData), Encoding.GetEncoding("Windows-1251"));
                MessageBox.Show(null, "Файл создан!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(null, "Ошибка при создании CSV файла:\r\n - " + ex.Message + "\r\n\r\nТрасировка ошибки:\r\n" + ex.StackTrace,
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //вспомогательные методы
        /// <summary> Получение инвентарного номера из листа. </summary>
        private string GetInventory(object[][] sheetData)
            => sheetData[1][1].ToString().Split('|')[2].Trim();

        /// <summary> Формирование содержимого CSV файла. </summary>
        private string BuildingContent(object[][] sheetData)
            => string.Join("\r\n", sheetData.Select(x => ParseData(x)).Where(x => x != null)) + "\r\n";

        /// <summary> Разбор и формирование строки для CSV. </summary>
        private string ParseData(object[] rowData)
        {
            string ConvertToString(object source)
               => source != null ? source.ToString() : string.Empty;
            int ConvertToInt(object source)
               => int.TryParse(ConvertToString(source), out int value) ? value : 0;

            var name = ConvertToString(rowData[5]).ToLower();
            if (name.Contains("лист"))
            {
                string position = ConvertToString(rowData[2]);
                string thickness = new string(name.SkipWhile(x => x < '0' || x > '9').TakeWhile(x => x >= '0' && x <= '9').ToArray());
                string count = (ConvertToInt(rowData[3]) + ConvertToInt(rowData[4])).ToString();
                string metal = ConvertToString(rowData[6]);
                return string.Join(";", new[] { position, count, thickness, metal });
            }
            else
                return null;
        }

        /// <summary> Получение данных Excel листа. </summary>
        private object[][] GetActiveSheetDataArray()
        {
            Range range = Globals.ThisAddIn.Application.ActiveSheet.UsedRange;

            return Enumerable.Range(1, range.Rows.Count).Select(row
                => Enumerable.Range(1, range.Columns.Count).Select(cell => (object)range.Value2[row, cell]).ToArray()
            ).ToArray();
        }
    }
}
