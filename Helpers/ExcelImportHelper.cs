using System.ComponentModel;
using System.ComponentModel.DataAnnotations.Schema;
using System.Reflection;
using OfficeOpenXml;
using Models;

public static class ExcelImportHelper
{
    public static List<ExcelPlayerModel> ImportPlayersFromExcel(string filePath, string sheetName)
    {
        var result = new List<ExcelPlayerModel>();

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var ws = package.Workbook.Worksheets[sheetName];
            if (ws == null) throw new Exception($"Sheet '{sheetName}' not found.");

            // Map column headers to property info
            var props = typeof(ExcelPlayerModel).GetProperties();
            var headerMap = new Dictionary<int, PropertyInfo>();
            for (int col = 1; col <= ws.Dimension.End.Column; col++)
            {
                var header = ws.Cells[1, col].Text.Trim();
                var prop = props.FirstOrDefault(p =>
                {
                    var attr = p.GetCustomAttribute<ColumnAttribute>();
                    return (attr != null && attr.Name == header) || p.Name == header;
                });
                if (prop != null)
                    headerMap[col] = prop;
            }

            // Read data rows
            for (int row = 2; row <= ws.Dimension.End.Row; row++)
            {
                var model = new ExcelPlayerModel
                {
                    Name = "",
                    IsPlayer = false,
                    IsSub = false,
                    Division = 0,
                    Team = "",
                    Handicap = 0
                };
                foreach (var kvp in headerMap)
                {
                    var col = kvp.Key;
                    var prop = kvp.Value;
                    var cellValue = ws.Cells[row, col].Text;

                    if (prop.PropertyType == typeof(string))
                        prop.SetValue(model, cellValue);
                    else if (prop.PropertyType == typeof(int))
                        prop.SetValue(model, int.TryParse(cellValue, out var i) ? i : 0);
                    else if (prop.PropertyType == typeof(bool))
                        prop.SetValue(model, cellValue == "1" || cellValue.Equals("true", StringComparison.OrdinalIgnoreCase));
                }
                result.Add(model);
            }
        }
        return result;
    }
}