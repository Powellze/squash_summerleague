using System.ComponentModel.DataAnnotations.Schema;
using Models;
using System.Reflection;
using ClosedXML.Excel;

namespace Helpers
{
    public class SheetData<T>
    {
        public required string SheetName { get; set; }
        public required List<T> Records { get; set; }
        public bool HasHeaderRow { get; set; } = true;
    }

    public static class FileHelper
    {
        //public static List<T> GetExcelFile<T>(string excelFilePath, string sheetName, bool hasHeaderRow = true, int headerRowNumber = 0)
        //{
        //    if (!File.Exists(excelFilePath))
        //    {
        //        throw new FileNotFoundException($"File is missing, tried to find file at: {excelFilePath}");
        //    }

        //    var excelFileMapper = new Ganss.Excel.ExcelMapper(excelFilePath);

        //    excelFileMapper.HeaderRow = hasHeaderRow;
        //    excelFileMapper.HeaderRowNumber = headerRowNumber;

        //    var records = excelFileMapper.Fetch<T>(sheetName);
        //    return records.ToList();
        //}

        public static List<ExcelPlayerModel> GetExcelFile(string filePath, string sheetName)
        {
            var result = new List<ExcelPlayerModel>();
            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet(sheetName);
                var props = typeof(ExcelPlayerModel).GetProperties();
                var headerMap = new Dictionary<int, PropertyInfo>();

                // Map headers to properties
                for (int col = 1; ; col++)
                {
                    var header = ws.Cell(1, col).GetString();
                    if (string.IsNullOrWhiteSpace(header)) break;
                    var prop = props.FirstOrDefault(p =>
                    {
                        var attr = p.GetCustomAttribute<ColumnAttribute>();
                        return (attr != null && attr.Name == header) || p.Name == header;
                    });
                    if (prop != null)
                        headerMap[col] = prop;
                }

                // Read data rows
                int row = 2;
                while (!ws.Cell(row, 1).IsEmpty())
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
                        var cellValue = ws.Cell(row, col).GetString();

                        if (prop.PropertyType == typeof(string))
                            prop.SetValue(model, cellValue);
                        else if (prop.PropertyType == typeof(int))
                            prop.SetValue(model, int.TryParse(cellValue, out var i) ? i : 0);
                        else if (prop.PropertyType == typeof(bool))
                            prop.SetValue(model, cellValue == "1" || cellValue.Equals("true", StringComparison.OrdinalIgnoreCase));
                    }
                    result.Add(model);
                    row++;
                }
            }
            return result;
        }

        public static string GetMicrosoftJsonFile(string jsonFilePath)
        {
            if (!File.Exists(jsonFilePath))
            {
                throw new FileNotFoundException($"File is missing, tried to find file at: {jsonFilePath}");
            }

            var jsonString = File.ReadAllText(jsonFilePath);


            return jsonString;

        }

        public static void WriteExcelFile<T>(IEnumerable<SheetData<T>> sheets, string excelFilePath)
        {
            var excelMapper = new Ganss.Excel.ExcelMapper();

            foreach (var sheet in sheets)
            {
                excelMapper.HeaderRow = sheet.HasHeaderRow;
                excelMapper.Save(excelFilePath, sheet.Records, sheet.SheetName);
            }
        }

        public static void WriteExcelFile<T>(List<T> records, string excelFilePath, string sheetName = "Sheet1", bool hasHeaderRow = true)
        {
            // Read the workbook into memory first
            byte[] fileBytes = File.ReadAllBytes(excelFilePath);
            using (var ms = new MemoryStream(fileBytes))
            {
                var workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(ms);

                int sheetIndex = workbook.GetSheetIndex(sheetName);
                if (sheetIndex != -1)
                {
                    workbook.RemoveSheetAt(sheetIndex);
                }

                // Save the workbook (without the old sheet) to a temp file
                var tempFile = Path.GetTempFileName();
                using (var tempFs = new FileStream(tempFile, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(tempFs);
                }

                // Use ExcelMapper to write the new data to the temp file
                var excelMapper = new Ganss.Excel.ExcelMapper(tempFile);
                excelMapper.HeaderRow = hasHeaderRow;
                excelMapper.Save(tempFile, records, sheetName);

                // Overwrite the original file with the updated temp file
                File.Copy(tempFile, excelFilePath, true);
                File.Delete(tempFile);
            }
        }

        public static class JsonCache
        {
            public enum JsonSerializer
            {
                Newtonsoft = 1,
                Microsoft = 2
            }

            public static async Task WriteObjectToJsonAsync<T>(string fullFilePath, T data, JsonSerializer serializer)
            {
                if (serializer == JsonSerializer.Microsoft)
                {
                    // Serialize the object to a JSON string
                    string jsonString = System.Text.Json.JsonSerializer.Serialize(data, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });

                    // Write the JSON string to the specified file
                    await File.WriteAllTextAsync(fullFilePath, jsonString);
                }
                else
                {
                    throw new NotImplementedException();
                }
            }

            public static async Task<T> ReadObjectFromJsonAsync<T>(string fullFilePath, JsonSerializer serializer) where T : new()
            {
                if (serializer == JsonSerializer.Microsoft)
                {
                    if (File.Exists(fullFilePath))
                    {
                        using (var reader = new StreamReader(fullFilePath))
                        {
                            string json = await reader.ReadToEndAsync();
                            var myObject = System.Text.Json.JsonSerializer.Deserialize<T>(json);
                            return myObject;
                        }
                    }
                    else
                    {
                        return new T();
                    }
                }
                else
                {
                    throw new NotImplementedException();
                }
            }

            public static async Task<T> ReadMultipleObjectsFromJsonAsync<T>(string directoryPath, string searchPattern, JsonSerializer serializer) where T : new()
            {
                if (serializer == JsonSerializer.Microsoft)
                {
                    // Get all files that match the search pattern
                    string[] files = Directory.GetFiles(directoryPath, searchPattern);

                    T mergedResult = new T();

                    if (typeof(T).IsGenericType && typeof(T).GetGenericTypeDefinition() == typeof(List<>))
                    {
                        var listType = typeof(T).GetGenericArguments();
                        var allItems = (IList<object>)Activator.CreateInstance(typeof(List<>).MakeGenericType(listType));

                        foreach (string file in files)
                        {
                            Console.WriteLine($"Reading file: {file}");
                            string content = await File.ReadAllTextAsync(file);
                            Console.WriteLine(content);
                            var myObject = System.Text.Json.JsonSerializer.Deserialize<T>(content);
                            if (myObject != null)
                            {
                                foreach (var item in (IEnumerable<object>)myObject)
                                {
                                    allItems.Add(item);
                                }
                            }
                        }

                        mergedResult = (T)allItems;
                    }

                    return mergedResult;
                }
                else
                {
                    throw new NotImplementedException();
                }
            }
        }
    }


    public static class JsonCache
    {
        public enum JsonSerializer
        {
            Newtonsoft = 1,
            Microsoft = 2
        }

        public static async Task WriteObjectToJsonAsync<T>(string fullFilePath, T data, JsonSerializer serializer)
        {
            if (serializer == JsonSerializer.Microsoft)
            {
                // Serialize the object to a JSON string
                string jsonString = System.Text.Json.JsonSerializer.Serialize(data, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });

                // Write the JSON string to the specified file
                await File.WriteAllTextAsync(fullFilePath, jsonString);
            }
            else
            {
                throw new NotImplementedException();
            }
        }

        public static async Task<T> ReadObjectFromJsonAsync<T>(string fullFilePath, JsonSerializer serializer) where T : new()
        {
            if (serializer == JsonSerializer.Microsoft)
            {
                if (File.Exists(fullFilePath))
                {
                    using (var reader = new StreamReader(fullFilePath))
                    {
                        string json = await reader.ReadToEndAsync();
                        var myObject = System.Text.Json.JsonSerializer.Deserialize<T>(json);
                        return myObject;
                    }
                }
                else
                {
                    return new T();
                }
            }
            else
            {
                throw new NotImplementedException();
            }
        }

        public static async Task<T> ReadMultipleObjectsFromJsonAsync<T>(string directoryPath, string searchPattern, JsonSerializer serializer) where T : new()
        {
            if (serializer == JsonSerializer.Microsoft)
            {
                // Get all files that match the search pattern
                string[] files = Directory.GetFiles(directoryPath, searchPattern);

                T mergedResult = new T();

                if (typeof(T).IsGenericType && typeof(T).GetGenericTypeDefinition() == typeof(List<>))
                {
                    var listType = typeof(T).GetGenericArguments();
                    var allItems = (IList<object>)Activator.CreateInstance(typeof(List<>).MakeGenericType(listType));

                    foreach (string file in files)
                    {
                        Console.WriteLine($"Reading file: {file}");
                        string content = await File.ReadAllTextAsync(file);
                        Console.WriteLine(content);
                        var myObject = System.Text.Json.JsonSerializer.Deserialize<T>(content);
                        if (myObject != null)
                        {
                            foreach (var item in (IEnumerable<object>)myObject)
                            {
                                allItems.Add(item);
                            }
                        }
                    }

                    mergedResult = (T)allItems;
                }

                return mergedResult;
            }
            else
            {
                throw new NotImplementedException();
            }
        }
    }
}