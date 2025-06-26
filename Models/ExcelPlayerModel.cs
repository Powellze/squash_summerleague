using System.ComponentModel.DataAnnotations.Schema;
using System.Reflection;

namespace Models;

public class ExcelPlayerModel
{
    public required string Name { get; set; }
    [Column("Player")] 
    public required bool IsPlayer { get; set; }
    [Column("Sub")] 
    public required bool IsSub { get; set; }
    public required int Division { get; set; }
    public required string Team { get; set; }
    public required int Handicap { get; set; }

    // Calculated properties with column mappings
    public string TotalPointsFormula { get; set; } = string.Empty;
    public string MatchesPlayedFormula { get; set; } = string.Empty;
    public string AverageFormula { get; set; } = string.Empty;

    /// <summary>
    /// Returns a dictionary mapping Excel column names (from [Column] or property name) to PropertyInfo.
    /// </summary>
    public static Dictionary<string, PropertyInfo> GetExcelColumnMap()
    {
        var props = typeof(ExcelPlayerModel).GetProperties();
        var map = new Dictionary<string, PropertyInfo>(StringComparer.OrdinalIgnoreCase);

        foreach (var prop in props)
        {
            var attr = prop.GetCustomAttribute<ColumnAttribute>();
            var colName = attr?.Name ?? prop.Name;
            map[colName] = prop;
        }
        return map;
    }
}