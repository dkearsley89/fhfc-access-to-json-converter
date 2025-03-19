using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;

namespace FhfcAccessToJsonConverter
{
    internal class Program
    {
        private static string? AccessDatabasePath;
        private static string? JsonFilePath;
        static void Main()
        {
            var sw = Stopwatch.StartNew();
            Console.WriteLine("Starting FhfcAccessToJsonConverter");
            var configuration = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
                .Build();
            AccessDatabasePath = configuration["AppSettings:AccessDatabasePath"] ?? "";
            JsonFilePath = configuration["AppSettings:JsonFilePath"] ?? "";
            RetrieveHomePageInfo();
            RetrieveSummaryRecords();
            RetrieveRecords();
            RetrieveMilestones();
            RetrievePlayers();
            RetrieveCoaches();
            RetrievePlayerIndividualInfo();
            RetrieveCoachIndividualInfo();
            sw.Stop();
            Console.WriteLine("Finished FhfcAccessToJsonConverter");
            Console.WriteLine("Time: " + sw.Elapsed.ToString(@"mm\:ss\.ff"));
            Console.Read();
        }
        private static void RetrieveHomePageInfo()
        {
            string[] homePageRecordsToDisplay = ["Most Senior Games", "Most A Grade Games", "Most A Grade Goals", "Most A Grade Premierships", "Most Open Women's Games", "Most Junior Games"];
            string jsonToReturn = "{\"records\":[";
            foreach (var record in homePageRecordsToDisplay)
            {
                SqlStatements.RecordSqlStatements.TryGetValue(record, out var dictionaryItem);
                if (dictionaryItem != null)
                {
                    var dt = QueryAccessDatabase(dictionaryItem[0].Replace("[NumberOfRecords]", "1"));
                    if (dt.Columns.Contains("Year"))
                    {
                        dt.Columns.Remove("Year");
                    }
                    if (dt.Columns.Contains("Round"))
                    {
                        dt.Columns.Remove("Round");
                    }
                    if (dt.Columns.Contains("Grade"))
                    {
                        dt.Columns.Remove("Grade");
                    }
                    jsonToReturn += "{\"name\":\"" + record + "\",\"label\":\"" + dictionaryItem[1] + "\",\"data\":" + JsonConvert.SerializeObject(dt) + "},";
                }
            }
            jsonToReturn = jsonToReturn.TrimEnd(',') + "]}";
            SaveJsonToFile(jsonToReturn, "homeRecords.json");
            SaveJsonToFile("{\"lastUpdated\":\"" + string.Format(DateTime.Now.ToString("MMMM d{0} yyyy"), GetDaySuffix(DateTime.Now.Day)) + "\"}", "lastUpdated.json");
        }
        private static void RetrieveSummaryRecords()
        {
            string jsonToReturn = "{\"records\":[";
            foreach (var item in SqlStatements.RecordSqlStatements)
            {
                var dt = QueryAccessDatabase(item.Value[0].Replace("[NumberOfRecords]", "5"));
                if (dt == null)
                {
                    return;
                }
                while (dt.Rows.Count > 5)
                {
                    dt.Rows.RemoveAt(dt.Rows.Count - 1);
                }
                if (dt.Columns.Contains("Year"))
                {
                    dt.Columns.Remove("Year");
                }
                if (dt.Columns.Contains("Round"))
                {
                    dt.Columns.Remove("Round");
                }
                if (dt.Columns.Contains("Grade"))
                {
                    dt.Columns.Remove("Grade");
                }
                jsonToReturn += "{\"name\":\"" + item.Key + "\",\"headers\":{\"c1\":\"Name\",\"c2\":\"" + item.Value[1] + "\"},\"data\":" + JsonConvert.SerializeObject(dt) + "},";
            }
            jsonToReturn = jsonToReturn.TrimEnd(',') + "]}";
            SaveJsonToFile(jsonToReturn, "records.json");
        }
        private static void RetrieveRecords()
        {
            foreach (var item in SqlStatements.RecordSqlStatements)
            {
                var dt = QueryAccessDatabase(item.Value[0].Replace("[NumberOfRecords]", "100"));
                if (dt == null)
                {
                    return;
                }
                while (dt.Rows.Count > 100)
                {
                    dt.Rows.RemoveAt(dt.Rows.Count - 1);
                }
                if (dt.Columns.Contains("Year"))
                {
                    dt.Columns.Remove("Year");
                }
                if (dt.Columns.Contains("Round"))
                {
                    dt.Columns.Remove("Round");
                }
                if (dt.Columns.Contains("Grade"))
                {
                    dt.Columns.Remove("Grade");
                }
                SaveJsonToFile("{\"name\":\"" + item.Key + "\",\"headers\":{\"c1\":\"Name\",\"c2\":\"" + item.Value[1] + "\"},\"data\":" + JsonConvert.SerializeObject(dt) + "}", item.Key + ".json");
            }
        }
        private static void RetrieveMilestones()
        {
            string jsonToReturn = "{\"milestones\":[";
            foreach (var item in SqlStatements.MilestoneSqlStatements)
            {
                var upcomingMilestones = QueryAccessDatabase(item.Value[0]);
                var recentMilestones = QueryAccessDatabase(item.Value[1]);
                if (upcomingMilestones.Rows.Count == 0 && recentMilestones.Rows.Count == 0)
                {
                    continue;
                }
                jsonToReturn += "{\"name\":\"" + item.Key + "\",\"grade\":\"" + item.Value[2] + "\",\"type\":\"" + item.Value[3] + "\",";
                bool upcomingMilestonesExist = false;
                if (upcomingMilestones.Rows.Count > 0)
                {
                    upcomingMilestonesExist = true;
                    jsonToReturn += "\"upcoming\":" + JsonConvert.SerializeObject(upcomingMilestones);
                }
                if (recentMilestones.Rows.Count > 0)
                {
                    if (upcomingMilestonesExist)
                    {
                        jsonToReturn += ",";
                    }
                    jsonToReturn += "\"recent\":" + JsonConvert.SerializeObject(recentMilestones);
                }
                jsonToReturn += "},";
            }
            jsonToReturn = string.Concat(jsonToReturn.AsSpan(0, jsonToReturn.Length - 1), "]}");
            SaveJsonToFile(jsonToReturn, "milestones.json");
        }
        private static void RetrievePlayers()
        {
            var dt = QueryAccessDatabase("SELECT DISTINCT[Games].Id AS id, FirstName & ' ' & LastName AS name, LastName, FirstName FROM [FHFC Membership List] INNER JOIN [Games] ON Games.Id = [FHFC Membership List].Id WHERE[Games].Id <> 1001 ORDER BY LastName, FirstName, [Games].Id");
            if (dt.Columns.Contains("FirstName"))
            {
                dt.Columns.Remove("FirstName");
            }
            if (dt.Columns.Contains("LastName"))
            {
                dt.Columns.Remove("LastName");
            }
            SaveJsonToFile("{\"players\":" + JsonConvert.SerializeObject(dt) + "}", "players.json");
        }
        private static void RetrieveCoaches()
        {
            var dt = QueryAccessDatabase("SELECT DISTINCT[GamesCoaches].Id AS id, FirstName & ' ' & LastName AS name, LastName, FirstName FROM [FHFC Membership List] INNER JOIN [GamesCoaches] ON GamesCoaches.Id = [FHFC Membership List].Id ORDER BY LastName, FirstName, [GamesCoaches].Id");
            if (dt.Columns.Contains("FirstName"))
            {
                dt.Columns.Remove("FirstName");
            }
            if (dt.Columns.Contains("LastName"))
            {
                dt.Columns.Remove("LastName");
            }
            SaveJsonToFile("{\"coaches\":" + JsonConvert.SerializeObject(dt) + "}", "coaches.json");
        }
        private static void RetrievePlayerIndividualInfo()
        {
            foreach (DataRow player in QueryAccessDatabase("SELECT DISTINCT Id FROM [Games] WHERE Id <> 1001").Rows)
            {
                var playerId = player["Id"].ToString();
                DataTable dt = QueryAccessDatabase(@"SELECT firstName, middleName, lastName, " +
                    "CINT((SELECT SUM(A) FROM [Games] WHERE ID = " + player["Id"] + ")) AS agrade, " +
                    "CINT((SELECT SUM(A) + SUM(B) + SUM(C) + SUM(OpenW) + SUM(Unknown_Senior) FROM [Games] WHERE ID = " + player["Id"] + ")) AS senior, " +
                    "CINT((SELECT SUM([18]) + SUM([17_5]) + SUM([17]) + SUM([17Girls]) + SUM([16]) + SUM([16_5sun]) + SUM([16Girls]) + SUM([16sun]) + SUM([15]) + SUM([15sun]) + SUM([14]) + SUM([14Girls]) + SUM([14sun]) + SUM([13]) + SUM([13sun]) + SUM(Unknown_Junior) FROM [Games] WHERE ID = " + player["Id"] + ")) AS junior, " +
                    "(SELECT MIN(Year) FROM [Games] WHERE ID = " + player["Id"] + ") AS minYear, " +
                    "(SELECT MAX(Year) FROM [Games] WHERE ID = " + player["Id"] + ") AS maxYear, " +
                    "(SELECT COUNT(Year) FROM [Games] WHERE ID = " + player["Id"] + ") AS seasons, " +
                    "CINT(IIf(IsNull((SELECT SUM(Goals) FROM[Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'A')), 0, (SELECT SUM(Goals) FROM[Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'A'))) AS aGradeGoals, " +
                    "CINT(IIf(IsNull((SELECT SUM(Goals) FROM[Games Details Per Round] WHERE ID = " + player["Id"] + " AND (Grade = 'A' OR Grade = 'B' OR Grade = 'C' OR Grade = 'OpenW'))), 0, (SELECT SUM(Goals) FROM[Games Details Per Round] WHERE ID = " + player["Id"] + " AND (Grade = 'A' OR Grade = 'B' OR Grade = 'C' OR Grade = 'OpenW')))) AS seniorGoals, " +
                    "CINT(IIf(IsNull((SELECT SUM(Goals) FROM[Games Details Per Round] WHERE ID = " + player["Id"] + " AND (Grade = 'U18' OR Grade = 'U17.5' OR Grade = 'U17Girls' OR Grade = 'U16' OR Grade = 'U16.5sun' OR Grade = 'U16Girls' OR Grade = 'U16sun' OR Grade = 'U15' OR Grade = 'U15sun' OR Grade = 'U14' OR Grade = 'U14sun' OR Grade = 'U14Girls' OR Grade = 'U13' OR Grade = 'U13sun'))), 0, (SELECT SUM(Goals) FROM[Games Details Per Round] WHERE ID = " + player["Id"] + " AND (Grade = 'U18' OR Grade = 'U17.5' OR Grade = 'U17Girls' OR Grade = 'U16' OR Grade = 'U16.5sun' OR Grade = 'U16Girls' OR Grade = 'U16sun' OR Grade = 'U15' OR Grade = 'U15sun' OR Grade = 'U14' OR Grade = 'U14sun' OR Grade = 'U14Girls' OR Grade = 'U13' OR Grade = 'U13sun')))) AS juniorGoals " +
                    "FROM [FHFC Membership List] " +
                    "WHERE ID = " + player["Id"]);
                string json = JsonConvert.SerializeObject(dt);
                json = json.Substring(1, json.Length - 3);
                dt = QueryAccessDatabase(@"SELECT Year AS y, " +
                    "A AS aGames, (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'A' AND [Games Details Per Round].Year = [Games].Year) AS aGoals, " +
                    "B AS bGames, (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'B' AND [Games Details Per Round].Year = [Games].Year) AS bGoals, " +
                    "C AS cGames, (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'C' AND [Games Details Per Round].Year = [Games].Year) AS cGoals, " +
                    "OpenW AS wGames, (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'OpenW' AND [Games Details Per Round].Year = [Games].Year) AS wGoals, " +
                    "Unknown_Senior AS uGames, (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'Unknown_Senior' AND [Games Details Per Round].Year = [Games].Year) AS uGoals, " +
                    "[18] AS u18Games, (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U18' AND [Games Details Per Round].Year = [Games].Year) AS u18Goals, " +
                    "[17_5] AS u175Games, (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U17.5' AND [Games Details Per Round].Year = [Games].Year) AS u175Goals, " +
                    "[17] AS u17Games, (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U17' AND [Games Details Per Round].Year = [Games].Year) AS u17Goals, " +
                    "[16] AS u16Games, (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U16' AND [Games Details Per Round].Year = [Games].Year) AS u16Goals, " +
                    "[15] AS u15Games, (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U15' AND [Games Details Per Round].Year = [Games].Year) AS u15Goals, " +
                    "[14] AS u14Games, (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U14' AND [Games Details Per Round].Year = [Games].Year) AS u14Goals, " +
                    "[13] AS u13Games, (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U13' AND [Games Details Per Round].Year = [Games].Year) AS u13Goals, " +
                    "[16_5sun] AS u165sunGames, (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U16.5sun' AND [Games Details Per Round].Year = [Games].Year) AS u165sunGoals, " +
                    "[16sun] AS u16sunGames, (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U16sun' AND [Games Details Per Round].Year = [Games].Year) AS u16sunGoals, " +
                    "[15sun] AS u15sunGames, (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U15sun' AND [Games Details Per Round].Year = [Games].Year) AS u15sunGoals, " +
                    "[14sun] AS u14sunGames, (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U14sun' AND [Games Details Per Round].Year = [Games].Year) AS u14sunGoals, " +
                    "[13sun] AS u13sunGames, (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U13sun' AND [Games Details Per Round].Year = [Games].Year) AS u13sunGoals, " +
                    "[17Girls] AS u17girlsGames, (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U17Girls' AND [Games Details Per Round].Year = [Games].Year) AS u17girlsGoals, " +
                    "[16Girls] AS u16girlsGames, (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U16Girls' AND [Games Details Per Round].Year = [Games].Year) AS u16girlsGoals, " +
                    "[14Girls] AS u14girlsGames, (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U14Girls' AND [Games Details Per Round].Year = [Games].Year) AS u14girlsGoals, " +
                    "Unknown_Junior AS ujGames, (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'Unknown_Junior' AND [Games Details Per Round].Year = [Games].Year) AS ujGoals " +
                    "FROM [Games] WHERE ID = " + player["Id"] + " " +
                    "UNION " +
                    "SELECT 'Total', (SELECT SUM(A) FROM [Games] WHERE ID = " + player["Id"] + "), (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'A'), " +
                    "(SELECT SUM(B) FROM [Games] WHERE ID = " + player["Id"] + "), (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'B'), " +
                    "(SELECT SUM(C) FROM [Games] WHERE ID = " + player["Id"] + "), (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'C'), " +
                    "(SELECT SUM(OpenW) FROM [Games] WHERE ID = " + player["Id"] + "), (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'OpenW'), " +
                    "(SELECT SUM(Unknown_Senior) FROM [Games] WHERE ID = " + player["Id"] + "), (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'Unknown_Senior'), " +
                    "(SELECT SUM([18]) FROM [Games] WHERE ID = " + player["Id"] + "), (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U18'), " +
                    "(SELECT SUM([17_5]) FROM [Games] WHERE ID = " + player["Id"] + "), (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U17.5'), " +
                    "(SELECT SUM([17]) FROM [Games] WHERE ID = " + player["Id"] + "), (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U17'), " +
                    "(SELECT SUM([16]) FROM [Games] WHERE ID = " + player["Id"] + "), (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U16'), " +
                    "(SELECT SUM([15]) FROM [Games] WHERE ID = " + player["Id"] + "), (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U15'), " +
                    "(SELECT SUM([14]) FROM [Games] WHERE ID = " + player["Id"] + "), (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U14'), " +
                    "(SELECT SUM([13]) FROM [Games] WHERE ID = " + player["Id"] + "), (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U13'), " +
                    "(SELECT SUM([16_5sun]) FROM [Games] WHERE ID = " + player["Id"] + "), (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U16.5sun'), " +
                    "(SELECT SUM([16sun]) FROM [Games] WHERE ID = " + player["Id"] + "), (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U16sun'), " +
                    "(SELECT SUM([15sun]) FROM [Games] WHERE ID = " + player["Id"] + "), (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U15sun'), " +
                    "(SELECT SUM([14sun]) FROM [Games] WHERE ID = " + player["Id"] + "), (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U14sun'), " +
                    "(SELECT SUM([13sun]) FROM [Games] WHERE ID = " + player["Id"] + "), (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U13sun'), " +
                    "(SELECT SUM([17Girls]) FROM [Games] WHERE ID = " + player["Id"] + "), (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U17Girls'), " +
                    "(SELECT SUM([16Girls]) FROM [Games] WHERE ID = " + player["Id"] + "), (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U16Girls'), " +
                    "(SELECT SUM([14Girls]) FROM [Games] WHERE ID = " + player["Id"] + "), (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U14Girls'), " +
                    "(SELECT SUM(Unknown_Junior) FROM [Games] WHERE ID = " + player["Id"] + "), (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'Unknown_Junior') " +
                    "FROM (SELECT COUNT(*) FROM [Games] WHERE 1 = 0) AS dual");
                string playerJson = ",\"years\":[";
                foreach (DataRow row in dt.Rows)
                {
                    playerJson += "{";
                    foreach (DataColumn column in row.Table.Columns)
                    {
                        if (column.ColumnName == "y")
                        {
                            playerJson += "\"" + column.ColumnName + "\":\"" + row[column] + "\",";
                        }
                        else if (!string.IsNullOrEmpty(row[column].ToString()) && row[column].ToString() != "0")
                        {
                            playerJson += "\"" + column.ColumnName + "\":" + row[column] + ",";
                        }
                    }
                    playerJson = playerJson.Substring(0, playerJson.Length - 1) + "},";
                }
                playerJson = string.Concat(playerJson.AsSpan(0, playerJson.Length - 1), "]}");
                SaveJsonToFile(json + playerJson, playerId + ".json");
            }
        }
        private static void RetrieveCoachIndividualInfo()
        {
            foreach (DataRow coach in QueryAccessDatabase("SELECT DISTINCT Id FROM GamesCoaches").Rows)
            {
                var coachId = coach["Id"].ToString();
                var dt = QueryAccessDatabase(@$"SELECT firstName, middleName, lastName,
                                               (SELECT SUM(A) FROM GamesCoaches WHERE ID = {coachId}) AS agrade,
                                               (SELECT SUM(A) + SUM(B) + SUM(C) + SUM(OpenW) FROM GamesCoaches WHERE ID = {coachId}) AS senior,
                                               (SELECT SUM([18]) + SUM([17_5]) + SUM([17]) + SUM([17Girls]) + SUM([16]) + SUM([16_5sun]) + SUM([16Girls]) + SUM([16sun]) + SUM([15]) + SUM([15sun]) + SUM([14]) + SUM([14Girls]) + SUM([14sun]) + SUM([13]) + SUM([13sun]) FROM GamesCoaches WHERE ID = {coachId}) AS junior,
                                               (SELECT MIN(Year) FROM GamesCoaches WHERE ID = {coachId}) AS minYear,
                                               (SELECT MAX(Year) FROM GamesCoaches WHERE ID = {coachId}) AS maxYear,
                                               (SELECT COUNT(Year) FROM GamesCoaches WHERE ID = {coachId}) AS seasons
                                               FROM [FHFC Membership List]
                                               WHERE ID = {coachId}");
                var json = JsonConvert.SerializeObject(dt);
                json = json.Substring(1, json.Length - 2);
                SaveJsonToFile(json, coachId + "-coach.json");
            }
        }
        private static void SaveJsonToFile(string json, string fileName)
        {
            using var sw = new StreamWriter(JsonFilePath + fileName);
            sw.Write(json);
        }
        private static DataTable QueryAccessDatabase(string sql)
        {
            OleDbConnection accessConnection;
            try
            {
                accessConnection = new OleDbConnection(AccessDatabasePath);
            }
            catch (Exception ex)
            {
                Console.Write("Error: Failed to create a database connection.\n{0}", ex.Message);
                return new DataTable();
            }
            try
            {
                var accessCommand = new OleDbCommand(sql, accessConnection);
                var dataAdapter = new OleDbDataAdapter(accessCommand);
                accessConnection.Open();
                var ds = new DataSet();
                dataAdapter.Fill(ds);
                return ds.Tables[0];
            }
            catch (Exception ex)
            {
                Console.Write("Error: Failed to retrieve the required data from the database.\n{0}", ex.Message);
                return new DataTable();
            }
            finally
            {
                accessConnection.Close();
            }
        }
        private static string GetDaySuffix(int day)
        {
            switch (day)
            {
                case 1:
                case 21:
                case 31:
                    return "st";
                case 2:
                case 22:
                    return "nd";
                case 3:
                case 23:
                    return "rd";
                default:
                    return "th";
            }
        }
    }
}