using ClosedXML.Excel;
using Helpers;
using Models;

namespace SpreadsheetGenerator
{
    internal class Program
    {
        // Add at the top of the Program class
        private static readonly Dictionary<string, int[]> SampleScores = new()
        {
            {"Tim Powell (A)", new[] {21, 17, 19}},
            {"Glen Scarah (B)", new[] {15, 21, 15}},
            {"Jordon Clark (C)", new[] {17, 15, 21}},
            {"Steve Stammers (D)", new[] {15, 19, 15}}
        };

        private static readonly Dictionary<string, int> RandomScores = SampleScores
            .ToDictionary(k => k.Key, v => v.Value[Random.Shared.Next(v.Value.Length)]);


        static void Main(string[] args)
        {
            var _createFixtureList = true; // Set to false if you don't want to create fixtures
            string excelFilePath = @"C:\Users\TimPowell\OneDrive - Herd Agency Limited\_META Shared Folders\Tim\Tim\squash\2025 Summer League\DL-Summer-League-2025.xlsx";

            var players = Helpers.FileHelper.GetExcelFile(excelFilePath, "Players").ToList();
            var playersByTeam = players.Where(w => w.IsPlayer).GroupBy(x => x.Team).ToList();

            // Convert to required types
            var teamNames = playersByTeam.Select(g => g.Key).ToList();
            var teamsDict = playersByTeam.ToDictionary(g => g.Key, g => g.ToList());


            if (_createFixtureList)
            {
                var teamFixtureList = CreateFixturesForTeams(teamNames, teamsDict);
                WriteStyledFixturesToExcel(teamFixtureList, excelFilePath, "Fixtures");
            }

            WriteDivisionScoreSheetsClosedXml(players, excelFilePath);
            //WriteHandicapSheet(players, excelFilePath, "Handicaps");
            //WriteLeagueTablesSheet(players, excelFilePath);
        }




        private static void WriteLeagueTablesSheet(List<ExcelPlayerModel> players, string filePath, string sheetName = "League Tables")
        {
            // First, populate the player stats from division scorecards
            using (var wb = new XLWorkbook(filePath))
            {
                foreach (var player in players.Where(p => p.IsPlayer))
                {
                    string scorecardSheet = $"Division {player.Division} Scorecard";
                    if (wb.Worksheets.TryGetWorksheet(scorecardSheet, out var ws))
                    {
                        // Find the player's column in the scorecard
                        int playerColumn = -1;
                        for (int col = 2; col <= ws.LastColumnUsed().ColumnNumber(); col++)
                        {
                            var headerCell = ws.Cell(1, col).Value.ToString();
                            if (headerCell.StartsWith(player.Name + " ("))
                            {
                                playerColumn = col;
                                break;
                            }
                        }

                        if (playerColumn != -1)
                        {
                            var lastRow = ws.LastRowUsed().RowNumber();
                            // Get Total Points from the "Total Points" row
                            var totalPoints = ws.Cell(lastRow - 2, playerColumn).FormulaA1;
                            var matchesPlayed = ws.Cell(lastRow - 1, playerColumn).FormulaA1;
                            var average = ws.Cell(lastRow, playerColumn).FormulaA1;

                            // Set formulas that reference the Division scorecard
                            player.TotalPointsFormula = $"'{scorecardSheet}'!{XLHelper.GetColumnLetterFromNumber(playerColumn)}{lastRow - 2}";
                            player.MatchesPlayedFormula = $"'{scorecardSheet}'!{XLHelper.GetColumnLetterFromNumber(playerColumn)}{lastRow - 1}";
                            player.AverageFormula = $"'{scorecardSheet}'!{XLHelper.GetColumnLetterFromNumber(playerColumn)}{lastRow}";
                        }
                    }
                }
            }

            // Now create the league tables sheet with formulas
            if (File.Exists(filePath))
            {
                using (var wb = new XLWorkbook(filePath))
                {
                    if (wb.Worksheets.Any(ws => ws.Name == sheetName))
                    {
                        wb.Worksheet(sheetName).Delete();
                        wb.Save();
                    }
                }
            }

            using (var wb = File.Exists(filePath) ? new XLWorkbook(filePath) : new XLWorkbook())
            {
                var ws = wb.Worksheets.Add(sheetName);
                int currentRow = 1;

                // Table 1: Team Positions
                ws.Cell(currentRow, 1).Value = "Team Positions";
                ws.Cell(currentRow, 1).Style.Font.Bold = true;
                currentRow++;

                var teamHeaders = new[] { "Position", "Team", "Total Points", "Played", "Average" };
                for (int i = 0; i < teamHeaders.Length; i++)
                    ws.Cell(currentRow, i + 1).Value = teamHeaders[i];

                var teamGroups = players
                    .Where(p => p.IsPlayer)
                    .GroupBy(p => p.Team)
                    .OrderBy(g => g.Key)
                    .ToList();

                // Inside the team groups loop in WriteLeagueTablesSheet:
                for (int i = 0; i < teamGroups.Count; i++)
                {
                    var row = currentRow + i + 1;
                    var team = teamGroups[i];
                    var teamName = team.Key;

                    ws.Cell(row, 1).Value = i + 1;
                    ws.Cell(row, 2).Value = teamName;

                    // Create single formula combining both divisions
                    ws.Cell(row, 3).FormulaA1 = $"=2*COUNTIFS('Division 1 Scorecard'!A2:A100,\"*({teamName})*\",'Division 1 Scorecard'!B2:Z100,\">15\")" +
                                                $"+2*COUNTIFS('Division 2 Scorecard'!A2:A100,\"*({teamName})*\",'Division 2 Scorecard'!B2:Z100,\">15\")";

                    // Count matches played
                    var matchesFormula = $"=COUNTIFS('Division 1 Scorecard'!A2:A100,\"*({teamName})*\",'Division 1 Scorecard'!B2:Z100,\"<>--\",'Division 1 Scorecard'!B2:Z100,\"<>\")" +
                                         $"+COUNTIFS('Division 2 Scorecard'!A2:A100,\"*({teamName})*\",'Division 2 Scorecard'!B2:Z100,\"<>--\",'Division 2 Scorecard'!B2:Z100,\"<>\")";

                    ws.Cell(row, 4).FormulaA1 = matchesFormula;

                    // Calculate average points per match
                    ws.Cell(row, 5).FormulaA1 = $"=IF(D{row}>0,C{row}/D{row},0)";
                    ws.Cell(row, 5).Style.NumberFormat.Format = "0.0";
                }

                var teamTableRange = ws.Range(currentRow, 1, currentRow + teamGroups.Count, 5);
                var teamTable = teamTableRange.CreateTable();
                teamTable.Theme = XLTableTheme.TableStyleMedium3;

                currentRow += teamGroups.Count + 4;

                // Add Division tables using modified WritePlayerTable method
                foreach (var division in new[] { 1, 2 })
                {
                    ws.Cell(currentRow, 1).Value = $"Division {division}";
                    ws.Cell(currentRow, 1).Style.Font.Bold = true;
                    currentRow++;

                    var divPlayers = players
                        .Where(p => p.IsPlayer && p.Division == division)
                        .ToList();

                    WritePlayerTableWithFormulas(ws, ref currentRow, divPlayers);

                    currentRow += 4;
                }

                // Format columns
                ws.Columns(1, 5).AdjustToContents();
                wb.SaveAs(filePath);
            }
        }

        private static void WritePlayerTableWithFormulas(IXLWorksheet ws, ref int startRow, List<ExcelPlayerModel> players)
        {
            var headers = new[] { "Position", "Name", "Total Points", "Played", "Average" };
            for (int i = 0; i < headers.Length; i++)
                ws.Cell(startRow, i + 1).Value = headers[i];

            for (int i = 0; i < players.Count; i++)
            {
                var row = startRow + i + 1;
                ws.Cell(row, 1).Value = i + 1;
                ws.Cell(row, 2).Value = players[i].Name;
                ws.Cell(row, 3).FormulaA1 = $"={players[i].TotalPointsFormula}";
                ws.Cell(row, 4).FormulaA1 = $"={players[i].MatchesPlayedFormula}";
                ws.Cell(row, 5).FormulaA1 = $"={players[i].AverageFormula}";
                ws.Cell(row, 5).Style.NumberFormat.Format = "0.0";
            }

            var tableRange = ws.Range(startRow, 1, startRow + players.Count, 5);
            var table = tableRange.CreateTable();
            table.Theme = XLTableTheme.TableStyleMedium3;

            startRow += players.Count + 1;
        }


        private static List<SimpleFixtureRowModel> CreateFixturesForTeams(List<string> teamNames, Dictionary<string, List<ExcelPlayerModel>> teamsDict)
        {
            bool randomizeFixtures = false;
            var fixtures = Fixtures.GenerateFixtures(teamNames, teamsDict, randomizeFixtures);

            // Now generate fixtures
            foreach (var fixture in fixtures.OrderBy(f => f.Week))
            {
                Console.WriteLine($"Week {fixture.Week}:");
                foreach (var team in fixture.Teams)
                {
                    Console.WriteLine($"  {team.Key} Players: {string.Join(", ", team.Value.Select(p => p.Name))}");
                }
            }

            // Set your league's start date
            DateTime startDate = new DateTime(2025, 6, 26); // Adjust as needed

            var simpleRows = new List<SimpleFixtureRowModel>();

            foreach (var fixture in fixtures.OrderBy(f => f.Week))
            {
                var byeTeams = fixture.Teams.Keys.Where(t => t.EndsWith("(BYE)")).ToList();
                var matchTeams = fixture.Teams.Keys.Where(t => !t.EndsWith("(BYE)")).ToList();
                var used = new HashSet<string>();

                // Output matches
                for (int i = 0; i < matchTeams.Count; i += 2)
                {
                    if (i + 1 < matchTeams.Count)
                    {
                        simpleRows.Add(new SimpleFixtureRowModel
                        {
                            Week = fixture.Week,
                            Date = startDate.AddDays(7 * (fixture.Week - 1)).ToString("dddd dd/MM/yyyy"),
                            Match = $"{matchTeams[i]} vs {matchTeams[i + 1]}"
                        });
                    }
                }

                // Output byes
                foreach (var bye in byeTeams)
                {
                    simpleRows.Add(new SimpleFixtureRowModel
                    {
                        Week = fixture.Week,
                        Date = startDate.AddDays(7 * (fixture.Week - 1)).ToString("dddd dd/MM/yyyy"),
                        Match = $"{bye.Replace(" (BYE)", "")} (BYE)"
                    });
                }
            }

            return simpleRows;
        }

        static void WriteHandicapSheet(List<ExcelPlayerModel> players, string filePath, string sheetName = "Handicaps")
        {
            // Remove the sheet if it exists
            if (File.Exists(filePath))
            {
                using (var wb = new XLWorkbook(filePath))
                {
                    if (wb.Worksheets.Any(ws => ws.Name == sheetName))
                    {
                        wb.Worksheet(sheetName).Delete();
                        wb.Save();
                    }
                }
            }

            using (var wb = File.Exists(filePath) ? new XLWorkbook(filePath) : new XLWorkbook())
            {
                var ws = wb.Worksheets.Add(sheetName);

                // Write header
                ws.Cell(1, 1).Value = "Name";
                ws.Cell(1, 2).Value = "Handicap";

                // Write data
                for (int i = 0; i < players.Count; i++)
                {
                    ws.Cell(i + 2, 1).Value = players[i].Name;
                    ws.Cell(i + 2, 2).Value = players[i].Handicap;
                }

                // Create table with Medium3 (lime) style
                var tableRange = ws.Range(1, 1, players.Count + 1, 2);
                var table = tableRange.CreateTable();
                table.Theme = XLTableTheme.TableStyleMedium3;

                // Auto-fit columns
                ws.Columns().AdjustToContents();

                wb.SaveAs(filePath);
            }
        }

        static void WriteStyledFixturesToExcel(List<SimpleFixtureRowModel> rows, string filePath, string sheetName)
        {
            // Remove the sheet if it exists
            if (File.Exists(filePath))
            {
                using (var wb = new XLWorkbook(filePath))
                {
                    if (wb.Worksheets.Any(ws => ws.Name == sheetName))
                    {
                        wb.Worksheet(sheetName).Delete();
                        wb.Save();
                    }
                }
            }

            // Open or create the workbook
            using (var wb = File.Exists(filePath) ? new XLWorkbook(filePath) : new XLWorkbook())
            {
                var ws = wb.Worksheets.Add(sheetName);

                // Write header
                var headers = new[] { "Week", "Date", "Match" };
                for (int i = 0; i < headers.Length; i++)
                {
                    ws.Cell(1, i + 1).Value = headers[i];
                }

                // Write data
                for (int i = 0; i < rows.Count; i++)
                {
                    ws.Cell(i + 2, 1).Value = rows[i].Week;
                    ws.Cell(i + 2, 2).Value = rows[i].Date;
                    ws.Cell(i + 2, 3).Value = rows[i].Match;
                }

                // Create table with Medium3 (lime) style
                var tableRange = ws.Range(1, 1, rows.Count + 1, headers.Length);
                var table = tableRange.CreateTable();
                table.Theme = XLTableTheme.TableStyleMedium3;

                // Auto-fit columns
                ws.Columns().AdjustToContents();

                // Optionally, set a max width for column B (Date)
                var dateCol = ws.Column(2);
                if (dateCol.Width > 20)
                    dateCol.Width = 20;

                wb.SaveAs(filePath);
            }
        }





        public static void WriteDivisionScoreSheetsClosedXml(List<ExcelPlayerModel> players, string filePath)
        {
            var divisions = players
                .Where(p => p.IsPlayer)
                .GroupBy(p => p.Division)
                .OrderBy(g => g.Key);

            // Remove old division sheets if they exist
            if (File.Exists(filePath))
            {
                using (var wb = new XLWorkbook(filePath))
                {
                    foreach (var division in divisions)
                    {
                        string sheetName = $"Division {division.Key} Scorecard";
                        if (wb.Worksheets.Any(ws => ws.Name == sheetName))
                        {
                            wb.Worksheet(sheetName).Delete();
                        }
                    }
                    wb.Save();
                }
            }

            var includeSampleScores = false; // Set to false if you don't want to include sample scores

            using (var wb = File.Exists(filePath) ? new XLWorkbook(filePath) : new XLWorkbook())
            {
                foreach (var division in divisions)
                {
                    string sheetName = $"Division {division.Key} Scorecard";
                    var ws = wb.Worksheets.Add(sheetName);

                    var playerList = division.OrderBy(p => p.Team).ToList();
                    int n = playerList.Count;

                    // Build data array
                    var headers = new string[n + 1];
                    headers[0] = "Players";
                    for (int i = 0; i < n; i++)
                        headers[i + 1] = $"{playerList[i].Name} ({playerList[i].Team})";

                    // Write header
                    for (int i = 0; i < headers.Length; i++)
                        ws.Cell(1, i + 1).Value = headers[i];

                    // Write grid
                    for (int row = 0; row < n; row++)
                    {
                        var rowPlayerName = $"{playerList[row].Name} ({playerList[row].Team})";

                        ws.Cell(row + 2, 1).Value = $"{playerList[row].Name} ({playerList[row].Team})";
                        for (int col = 0; col < n; col++)
                        {
                            var cell = ws.Cell(row + 2, col + 2);
                            if (row == col)
                            {
                                cell.Value = "--";
                                cell.Style.Fill.BackgroundColor = XLColor.Gray;
                                cell.Style.Font.FontColor = XLColor.White;
                                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                            }
                            else if (includeSampleScores)
                            {
                                var colPlayerName = $"{playerList[col].Name} ({playerList[col].Team})";

                                // Check if we have a score for this player
                                if (RandomScores.TryGetValue(rowPlayerName, out int score))
                                {
                                    cell.Value = score;

                                    // Style based on win/loss
                                    if (RandomScores.TryGetValue(colPlayerName, out int opponentScore))
                                    {
                                        var isWin = score > opponentScore;
                                        cell.Style.Fill.BackgroundColor = isWin ? XLColor.LightGreen : XLColor.LightPink;
                                        cell.Style.Font.Bold = isWin;
                                    }
                                }
                            }
                            else
                            {
                                cell.Value = "";
                            }
                        }
                    }

                    var colourBackground = XLColor.Gray;
                    var colourFont = XLColor.White;

                    // Create table with Medium3 (lime) style
                    var tableRange = ws.Range(1, 1, n + 1, n + 1);
                    var table = tableRange.CreateTable();
                    table.Theme = XLTableTheme.TableStyleMedium3;

                    // Add "Total Points" row (below table)
                    ws.Cell(n + 2, 1).Value = "Total Points";
                    ws.Cell(n + 2, 1).Style.Font.Bold = true;
                    ws.Cell(n + 2, 1).Style.Fill.BackgroundColor = colourBackground;
                    ws.Cell(n + 2, 1).Style.Font.FontColor = colourFont;
                    for (int col = 1; col <= n; col++)
                    {
                        // B2:B{n+1}, C2:C{n+1}, etc.
                        var colLetter = XLHelper.GetColumnLetterFromNumber(col + 1);
                        ws.Cell(n + 2, col + 1).FormulaA1 = $"SUMIF({colLetter}2:{colLetter}{n + 1},\">0\")";
                        ws.Cell(n + 2, col + 1).Style.Fill.BackgroundColor = colourBackground;
                        ws.Cell(n + 2, col + 1).Style.Font.FontColor = colourFont;

                        ws.Cell(n + 2, col + 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        ws.Cell(n + 2, col + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    }

                    // Add "Played Matches" row
                    ws.Cell(n + 3, 1).Value = "Played Matches";
                    ws.Cell(n + 3, 1).Style.Font.Bold = true;
                    ws.Cell(n + 3, 1).Style.Fill.BackgroundColor = colourBackground;
                    ws.Cell(n + 3, 1).Style.Font.FontColor = colourFont;
                    for (int col = 1; col <= n; col++)
                    {
                        var colLetter = XLHelper.GetColumnLetterFromNumber(col + 1);
                        ws.Cell(n + 3, col + 1).FormulaA1 = $"COUNTIF({colLetter}2:{colLetter}{n + 1},\">0\")";
                        ws.Cell(n + 3, col + 1).Style.Fill.BackgroundColor = colourBackground;
                        ws.Cell(n + 3, col + 1).Style.Font.FontColor = colourFont;

                        ws.Cell(n + 3, col + 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        ws.Cell(n + 3, col + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    }

                    // Add "AVG Points" row (1 decimal place)
                    ws.Cell(n + 4, 1).Value = "Average Points";
                    ws.Cell(n + 4, 1).Style.Fill.BackgroundColor = colourBackground;
                    ws.Cell(n + 4, 1).Style.Font.FontColor = colourFont;
                    ws.Cell(n + 4, 1).Style.Font.Bold = true;

                    for (int col = 1; col <= n; col++)
                    {
                        var totalCell = ws.Cell(n + 2, col + 1).Address.ToStringRelative();
                        var playedCell = ws.Cell(n + 3, col + 1).Address.ToStringRelative();
                        var avgCell = ws.Cell(n + 4, col + 1);
                        avgCell.FormulaA1 = $"IFERROR({totalCell}/{playedCell}, \"0\")";
                        avgCell.Style.NumberFormat.Format = "0.0";
                        avgCell.Style.Fill.BackgroundColor = colourBackground;
                        avgCell.Style.Font.FontColor = colourFont;

                        avgCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        avgCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    }

                    // Set column widths and row heights
                    ws.Columns().AdjustToContents();
                    for (int row = 1; row <= ws.LastRowUsed().RowNumber(); row++)
                        ws.Row(row).Height = 40;
                }

                wb.SaveAs(filePath);
            }
        }


    }
}
