﻿using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;

namespace FhfcAccessToJsonConverter
{
    internal class Program
    {
        private static readonly Dictionary<string, string[]> RecordSqlStatements = new Dictionary<string, string[]>
        {
            {"Most Senior Games", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT((SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior))) AS a FROM Games INNER JOIN [FHFC Membership List] ON Games.ID = [FHFC Membership List].ID GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior) > 0 ORDER BY (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Games" } },
            {"Most Club Games", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT((SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14Girls])+SUM(Games.[14sun])+SUM(Games.[13])+SUM(Games.unknown_Senior)+SUM(Games.unknown_Junior))) AS a FROM Games INNER JOIN [FHFC Membership List] ON Games.ID = [FHFC Membership List].ID GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName ORDER BY (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14Girls])+SUM(Games.[14sun])+SUM(Games.[13])+SUM(Games.unknown_Senior)+SUM(Games.unknown_Junior)) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Games" } },
            {"Most A Grade Games", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Games.a)) AS a FROM Games INNER JOIN [FHFC Membership List] ON Games.ID = [FHFC Membership List].ID GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Games.a) > 0 ORDER BY SUM(Games.a) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Games" } },

            {"Most B Grade Games", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Games.b)) AS a FROM Games INNER JOIN [FHFC Membership List] ON Games.ID = [FHFC Membership List].ID GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Games.b) > 0 ORDER BY SUM(Games.b) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Games" } },
            {"Most C Grade Games", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Games.c)) AS a FROM Games INNER JOIN [FHFC Membership List] ON Games.ID = [FHFC Membership List].ID GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Games.c) > 0 ORDER BY SUM(Games.c) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Games" } },
            {"Most Open Women's Games", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Games.openW)) AS a FROM Games INNER JOIN [FHFC Membership List] ON Games.ID = [FHFC Membership List].ID GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Games.openW) > 0 ORDER BY SUM(Games.openW) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Games" } },

            {"Most Senior Goals", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'A' OR [Games Details Per Round].Grade = 'B' OR [Games Details Per Round].Grade = 'C' OR [Games Details Per Round].Grade = 'OpenW' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most Club Goals", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most A Grade Goals", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'A' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },

            {"Most B Grade Goals", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'B' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most C Grade Goals", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'C' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most Open Women's Goals", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'OpenW' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },

            {"Avg Goals Per Game - Senior", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, FORMAT(AVG(Goals),\"0.,00\") AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'A' OR [Games Details Per Round].Grade = 'B' OR [Games Details Per Round].Grade = 'C' OR [Games Details Per Round].Grade = 'OpenW' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY AVG(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Avg Goals Per Game - Club", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, FORMAT(AVG(Goals),\"0.,00\") AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY AVG(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Avg Goals Per Game - A Grade", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, FORMAT(AVG(Goals),\"0.,00\") AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'A' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING AVG(Goals) > 0 ORDER BY AVG(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },

            {"Avg Goals Per Game - B Grade", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, FORMAT(AVG(Goals),\"0.,00\") AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'B' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY AVG(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Avg Goals Per Game - C Grade", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, FORMAT(AVG(Goals),\"0.,00\") AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'C' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY AVG(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Avg Goals Per Game - Open Women's", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, FORMAT(AVG(Goals),\"0.,00\") AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'OpenW' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY AVG(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },

            {"Most Goals in a Season - Senior", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'A' OR [Games Details Per Round].Grade = 'B' OR [Games Details Per Round].Grade = 'C' OR [Games Details Per Round].Grade = 'OpenW' GROUP BY [FHFC Membership List].id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most Goals in a Season - Club", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID GROUP BY [FHFC Membership List].id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most Goals in a Season - A Grade", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'A' GROUP BY [FHFC Membership List].id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },

            {"Most Goals in a Season - B Grade", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'B' GROUP BY [FHFC Membership List].id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most Goals in a Season - C Grade", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'C' GROUP BY [FHFC Membership List].id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most Goals in a Season - Open Women's", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'OpenW' GROUP BY [FHFC Membership List].id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },

            {"Most Goals in a Game - Senior", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, Year, Round, Grade, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, Goals AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'A' OR [Games Details Per Round].Grade = 'B' OR [Games Details Per Round].Grade = 'C' OR [Games Details Per Round].Grade = 'OpenW' GROUP BY [FHFC Membership List].id, Year, Round, Grade, Goals, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING Goals > 0 ORDER BY Goals DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most Goals in a Game - Club", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, Year, Round, Grade, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, Goals AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID GROUP BY [FHFC Membership List].id, Year, Round, Grade, Goals, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING Goals > 0 ORDER BY Goals DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most Goals in a Game - A Grade", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, Year, Round, Grade, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, Goals AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'A' GROUP BY [FHFC Membership List].id, Year, Round, Grade, Goals, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING Goals > 0 ORDER BY Goals DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },

            {"Most Goals in a Game - B Grade", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, Year, Round, Grade, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, Goals AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'B' GROUP BY [FHFC Membership List].id, Year, Round, Grade, Goals, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING Goals > 0 ORDER BY Goals DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most Goals in a Game - C Grade", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, Year, Round, Grade, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, Goals AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'C' GROUP BY [FHFC Membership List].id, Year, Round, Grade, Goals, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING Goals > 0 ORDER BY Goals DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most Goals in a Game - Open Women's", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, Year, Round, Grade, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, Goals AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'OpenW' GROUP BY [FHFC Membership List].id, Year, Round, Grade, Goals, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING Goals > 0 ORDER BY Goals DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },

            {"Most Junior Games", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT((SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14Girls])+SUM(Games.[14sun])+SUM(Games.[13])+SUM(Games.unknown_Junior))) AS a FROM Games INNER JOIN [FHFC Membership List] ON Games.ID = [FHFC Membership List].ID GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14Girls])+SUM(Games.[14sun])+SUM(Games.[13])+SUM(Games.unknown_Junior) > 0 ORDER BY (SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14Girls])+SUM(Games.[14sun])+SUM(Games.[13])+SUM(Games.unknown_Junior)) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Games" } },
            {"Most Junior Goals", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'U18' OR [Games Details Per Round].Grade = 'U17.5' OR [Games Details Per Round].Grade = 'U16' OR [Games Details Per Round].Grade = 'U16sun' OR [Games Details Per Round].Grade = 'U15' OR [Games Details Per Round].Grade = 'U15sun' OR [Games Details Per Round].Grade = 'U14' OR [Games Details Per Round].Grade = 'U14sun' OR [Games Details Per Round].Grade = 'U14Girls' OR [Games Details Per Round].Grade = 'U13' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Avg Goals Per Game - Junior", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, FORMAT(AVG(Goals),\"0.,00\") AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'U18' OR [Games Details Per Round].Grade = 'U17.5' OR [Games Details Per Round].Grade = 'U16' OR [Games Details Per Round].Grade = 'U16sun' OR [Games Details Per Round].Grade = 'U15' OR [Games Details Per Round].Grade = 'U15sun' OR [Games Details Per Round].Grade = 'U14' OR [Games Details Per Round].Grade = 'U14sun' OR [Games Details Per Round].Grade = 'U14Girls' OR [Games Details Per Round].Grade = 'U13' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY AVG(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },

            {"Most Goals in a Season - Junior", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'U18' OR [Games Details Per Round].Grade = 'U17.5' OR [Games Details Per Round].Grade = 'U16' OR [Games Details Per Round].Grade = 'U16sun' OR [Games Details Per Round].Grade = 'U15' OR [Games Details Per Round].Grade = 'U15sun' OR [Games Details Per Round].Grade = 'U14' OR [Games Details Per Round].Grade = 'U14sun' OR [Games Details Per Round].Grade = 'U14Girls' OR [Games Details Per Round].Grade = 'U13' GROUP BY [FHFC Membership List].id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most Goals in a Game - Junior", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, Year, Round, Grade, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, Goals AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'U18' OR [Games Details Per Round].Grade = 'U17.5' OR [Games Details Per Round].Grade = 'U16' OR [Games Details Per Round].Grade = 'U16sun' OR [Games Details Per Round].Grade = 'U15' OR [Games Details Per Round].Grade = 'U15sun' OR [Games Details Per Round].Grade = 'U14' OR [Games Details Per Round].Grade = 'U14sun' OR [Games Details Per Round].Grade = 'U14Girls' OR [Games Details Per Round].Grade = 'U13' GROUP BY [FHFC Membership List].id, Year, Round, Grade, Goals, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING Goals > 0 ORDER BY Goals DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most Seasons Played", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, COUNT(Games.Id) AS a FROM Games INNER JOIN [FHFC Membership List] ON Games.ID = [FHFC Membership List].ID GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName ORDER BY COUNT(Games.Id) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Seasons" } }
        };
        //private readonly static Dictionary<string, string> _milestoneSqlStatements = new Dictionary<string, string>()
        //{
        //    {"A Grade Games", "SELECT [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, SUM(Games.a) AS a FROM Games INNER JOIN [FHFC Membership List] ON Games.ID = [FHFC Membership List].ID WHERE [FHFC Membership List].currentlyActive = true GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING (SUM(Games.a) BETWEEN 45 AND 49 OR SUM(Games.a) BETWEEN 90 AND 99 OR SUM(Games.a) BETWEEN 145 AND 149 OR SUM(Games.a) BETWEEN 190 AND 199 OR SUM(Games.a) BETWEEN 245 AND 249 OR SUM(Games.a) BETWEEN 290 AND 299) AND MAX(Games.Year) > Year(Now())-2 ORDER BY SUM(Games.a) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName"},
        //    {"B Grade Games", "SELECT [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, SUM(Games.b) AS a FROM Games INNER JOIN [FHFC Membership List] ON Games.ID = [FHFC Membership List].ID WHERE [FHFC Membership List].currentlyActive = true GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING (SUM(Games.b) BETWEEN 45 AND 49 OR SUM(Games.b) BETWEEN 90 AND 99 OR SUM(Games.b) BETWEEN 145 AND 149 OR SUM(Games.b) BETWEEN 190 AND 199 OR SUM(Games.b) BETWEEN 245 AND 249 OR SUM(Games.b) BETWEEN 290 AND 299) AND MAX(Games.Year) > Year(Now())-2 ORDER BY SUM(Games.b) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName"},
        //    {"C Grade Games", "SELECT [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, SUM(Games.c) AS a FROM Games INNER JOIN [FHFC Membership List] ON Games.ID = [FHFC Membership List].ID WHERE [FHFC Membership List].currentlyActive = true GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING (SUM(Games.c) BETWEEN 45 AND 49 OR SUM(Games.c) BETWEEN 90 AND 99 OR SUM(Games.c) BETWEEN 145 AND 149 OR SUM(Games.c) BETWEEN 190 AND 199 OR SUM(Games.c) BETWEEN 245 AND 249 OR SUM(Games.c) BETWEEN 290 AND 299) AND MAX(Games.Year) > Year(Now())-2 ORDER BY SUM(Games.c) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName"},
        //    {"Open Women Games", "SELECT [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, SUM(Games.openW) AS a FROM Games INNER JOIN [FHFC Membership List] ON Games.ID = [FHFC Membership List].ID WHERE [FHFC Membership List].currentlyActive = true GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING (SUM(Games.openW) BETWEEN 45 AND 49 OR SUM(Games.openW) BETWEEN 90 AND 99 OR SUM(Games.openW) BETWEEN 145 AND 149 OR SUM(Games.openW) BETWEEN 190 AND 199 OR SUM(Games.openW) BETWEEN 245 AND 249 OR SUM(Games.openW) BETWEEN 290 AND 299) AND MAX(Games.Year) > Year(Now())-2 ORDER BY SUM(Games.openW) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName"},
        //    {"Senior Games", "SELECT [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)) AS a FROM Games INNER JOIN [FHFC Membership List] ON Games.ID = [FHFC Membership List].ID WHERE [FHFC Membership List].currentlyActive = true GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING ((SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)) BETWEEN 45 AND 49 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)) BETWEEN 90 AND 99 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)) BETWEEN 145 AND 149 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)) BETWEEN 190 AND 199 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)) BETWEEN 245 AND 249 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)) BETWEEN 290 AND 299 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)) BETWEEN 345 AND 349OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)) BETWEEN 390 AND 399OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)) BETWEEN 445 AND 449 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)) BETWEEN 490 AND 499) AND MAX(Games.Year) > Year(Now())-2 ORDER BY (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName"},
        //    {"Junior Games", "SELECT [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, (SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) AS a FROM Games INNER JOIN [FHFC Membership List] ON Games.ID = [FHFC Membership List].ID WHERE [FHFC Membership List].currentlyActive = true AND [FHFC Membership List].membershipStatus = 'jp' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING ((SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 45 AND 49 OR (SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 90 AND 99 OR (SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 145 AND 149 OR (SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 190 AND 199) AND MAX(Games.Year) > Year(Now())-2 ORDER BY (SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName"},
        //    {"Club Games", "SELECT [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) AS a FROM Games INNER JOIN [FHFC Membership List] ON Games.ID = [FHFC Membership List].ID WHERE [FHFC Membership List].currentlyActive = true GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING ((SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 45 AND 49 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 90 AND 99 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 145 AND 149 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 190 AND 199 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 245 AND 249 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 290 AND 299 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 345 AND 349OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 390 AND 399OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 445 AND 449 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 490 AND 499) AND MAX(Games.Year) > Year(Now())-2 ORDER BY (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName"},
        //    {"A Grade Goals", "SELECT [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, SUM(Goals) AS a FROM [Games Details per round] INNER JOIN [FHFC Membership List] ON [Games Details per round].ID = [FHFC Membership List].ID WHERE [FHFC Membership List].currentlyActive = true AND [Games Details per round].Grade = 'A' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING (SUM(Goals) BETWEEN 80 AND 99 OR  SUM(Goals) BETWEEN 180 AND 199 OR SUM(Goals) BETWEEN 280 AND 299 OR SUM(Goals) BETWEEN 380 AND 399 OR SUM(Goals) BETWEEN 480 AND 499 OR SUM(Goals) BETWEEN 580 AND 599 OR SUM(Goals) BETWEEN 680 AND 699 OR SUM(Goals) BETWEEN 780 AND 799 OR SUM(Goals) BETWEEN 880 AND 899 OR SUM(Goals) BETWEEN 980 AND 999) AND MAX(Year) > Year(Now())-2 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName"},
        //    {"Senior Goals", "SELECT [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, SUM(Goals) AS a FROM [Games Details per round] INNER JOIN [FHFC Membership List] ON [Games Details per round].ID = [FHFC Membership List].ID WHERE [FHFC Membership List].currentlyActive = true AND ([Games Details per round].Grade = 'A' OR [Games Details per round].Grade = 'B' OR [Games Details per round].Grade = 'C' OR [Games Details per round].Grade = 'OpenW') GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING (SUM(Goals) BETWEEN 80 AND 99 OR  SUM(Goals) BETWEEN 180 AND 199 OR SUM(Goals) BETWEEN 280 AND 299 OR SUM(Goals) BETWEEN 380 AND 399 OR SUM(Goals) BETWEEN 480 AND 499 OR SUM(Goals) BETWEEN 580 AND 599 OR SUM(Goals) BETWEEN 680 AND 699 OR SUM(Goals) BETWEEN 780 AND 799 OR SUM(Goals) BETWEEN 880 AND 899 OR SUM(Goals) BETWEEN 980 AND 999) AND MAX(Year) > Year(Now())-2 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName"},
        //    {"Junior Goals", "SELECT [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, SUM(Goals) AS a FROM [Games Details per round] INNER JOIN [FHFC Membership List] ON [Games Details per round].ID = [FHFC Membership List].ID WHERE [FHFC Membership List].currentlyActive = true AND [FHFC Membership List].membershipStatus = 'jp' AND ([Games Details per round].Grade = 'U18' OR [Games Details per round].Grade = 'U17.5' OR [Games Details per round].Grade = 'U17' OR [Games Details per round].Grade = 'U16' OR [Games Details per round].Grade = 'U16sun' OR [Games Details per round].Grade = 'U15' OR [Games Details per round].Grade = 'U15sun' OR [Games Details per round].Grade = 'U14' OR [Games Details per round].Grade = 'U14sun' OR [Games Details per round].Grade = 'U14Girls' OR [Games Details per round].Grade = 'U13') GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING (SUM(Goals) BETWEEN 80 AND 99 OR  SUM(Goals) BETWEEN 180 AND 199 OR SUM(Goals) BETWEEN 280 AND 299 OR SUM(Goals) BETWEEN 380 AND 399 OR SUM(Goals) BETWEEN 480 AND 499 OR SUM(Goals) BETWEEN 580 AND 599 OR SUM(Goals) BETWEEN 680 AND 699 OR SUM(Goals) BETWEEN 780 AND 799 OR SUM(Goals) BETWEEN 880 AND 899 OR SUM(Goals) BETWEEN 980 AND 999) AND MAX(Year) > Year(Now())-2 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName"},
        //    {"Club Goals", "SELECT [FHFC Membership List].id AS Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, SUM(Goals) AS a FROM [Games Details per round] INNER JOIN [FHFC Membership List] ON [Games Details per round].ID = [FHFC Membership List].ID WHERE [FHFC Membership List].currentlyActive = true GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING (SUM(Goals) BETWEEN 80 AND 99 OR  SUM(Goals) BETWEEN 180 AND 199 OR SUM(Goals) BETWEEN 280 AND 299 OR SUM(Goals) BETWEEN 380 AND 399 OR SUM(Goals) BETWEEN 480 AND 499 OR SUM(Goals) BETWEEN 580 AND 599 OR SUM(Goals) BETWEEN 680 AND 699 OR SUM(Goals) BETWEEN 780 AND 799 OR SUM(Goals) BETWEEN 880 AND 899 OR SUM(Goals) BETWEEN 980 AND 999) AND MAX(Year) > Year(Now())-2 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName"}
        //};
        private static OleDbConnection _accessConnection;
        private static void Main()
        {
            Console.WriteLine("Starting FhfcAccessToJsonConverter");
            RetrieveHomePageInfo();
            RetrieveSummaryRecords();
            RetrieveRecords();
            //RetrieveMilestones();
            RetrievePlayers();
            RetrievePlayerIndividualInfo();
            Console.WriteLine("Finishing FhfcAccessToJsonConverter");
            Console.Read();
        }
        private static void RetrieveHomePageInfo()
        {
            string[] homePageRecordsToDisplay = new[] { "Most Senior Games", "Most A Grade Games", "Most Junior Games", "Most Junior Goals", "Most Goals in a Season - Senior", "Most A Grade Goals" };
            string jsonToReturn = "{\"records\":[";
            foreach (KeyValuePair<string, string[]> item in RecordSqlStatements)
            {
                if (homePageRecordsToDisplay.Contains(item.Key))
                {
                    DataTable dt = QueryAccessDatabase(item.Value[0].Replace("[NumberOfRecords]", "1"));
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
                    jsonToReturn += "{\"name\":\"" + item.Key + "\",\"label\":\"" + item.Value[1] + "\",\"data\":" + Newtonsoft.Json.JsonConvert.SerializeObject(dt) + "},";
                }
            }
            jsonToReturn = jsonToReturn.TrimEnd(',') + "]}";
            SaveJsonToFile(jsonToReturn, "homeRecords.json");
            SaveJsonToFile("{\"lastUpdated\":\"" + String.Format(DateTime.Now.ToString("MMMM d{0} yyyy"), GetDaySuffix(DateTime.Now.Day)) + "\"}", "lastUpdated.json");
        }
        private static void RetrieveSummaryRecords()
        {
            string jsonToReturn = "{\"records\":[";
            foreach (KeyValuePair<string, string[]> item in RecordSqlStatements)
            {
                DataTable dt = QueryAccessDatabase(item.Value[0].Replace("[NumberOfRecords]", "5"));
                if (dt.Columns.Contains("Id"))
                {
                    dt.Columns.Remove("Id");
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
                jsonToReturn += "{\"name\":\"" + item.Key + "\",\"headers\":{\"c1\":\"Name\",\"c2\":\"" + item.Value[1] + "\"},\"data\":" + Newtonsoft.Json.JsonConvert.SerializeObject(dt) + "},";
            }
            jsonToReturn = jsonToReturn.TrimEnd(',') + "]}";
            SaveJsonToFile(jsonToReturn, "records.json");
        }
        private static void RetrieveRecords()
        {
            foreach (KeyValuePair<string, string[]> item in RecordSqlStatements)
            {
                DataTable dt = QueryAccessDatabase(item.Value[0].Replace("[NumberOfRecords]", "100"));
                if (dt.Columns.Contains("Id"))
                {
                    dt.Columns.Remove("Id");
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
                SaveJsonToFile("{\"name\":\"" + item.Key + "\",\"headers\":{\"c1\":\"Name\",\"c2\":\"" + item.Value[1] + "\"},\"data\":" + Newtonsoft.Json.JsonConvert.SerializeObject(dt) + "}", item.Key + ".json");
            }
        }
        private static void RetrieveMilestones(string fileName)
        {
            //string jsonToReturn = "{\"milestones\":[";
            //foreach (KeyValuePair<string, string> item in _milestoneSqlStatements)
            //{
            //    DataTable dt = QueryAccessDatabase(item.Value);
            //    dt.Columns.Remove("Id");
            //    string output = Newtonsoft.Json.JsonConvert.SerializeObject(dt);
            //    switch (item.Key)
            //    {
            //        case "A Grade Games":
            //            jsonToReturn += "{\"name\":\"A Grade Games\",\"headers\":{\"c1\":\"Name\",\"c2\":\"Games\"},\"data\":" + output + "},";
            //            break;
            //        case "B Grade Games":
            //            jsonToReturn += "{\"name\":\"B Grade Games\",\"headers\":{\"c1\":\"Name\",\"c2\":\"Games\"},\"data\":" + output + "},";
            //            break;
            //        case "C Grade Games":
            //            jsonToReturn += "{\"name\":\"C Grade Games\",\"headers\":{\"c1\":\"Name\",\"c2\":\"Games\"},\"data\":" + output + "},";
            //            break;
            //        case "Open Women Games":
            //            jsonToReturn += "{\"name\":\"Open Women Games\",\"headers\":{\"c1\":\"Name\",\"c2\":\"Games\"},\"data\":" + output + "},";
            //            break;
            //        case "Senior Games":
            //            jsonToReturn += "{\"name\":\"Senior Games\",\"headers\":{\"c1\":\"Name\",\"c2\":\"Games\"},\"data\":" + output + "},";
            //            break;
            //        case "Junior Games":
            //            jsonToReturn += "{\"name\":\"Junior Games\",\"headers\":{\"c1\":\"Name\",\"c2\":\"Games\"},\"data\":" + output + "},";
            //            break;
            //        case "Club Games":
            //            jsonToReturn += "{\"name\":\"Club Games\",\"headers\":{\"c1\":\"Name\",\"c2\":\"Games\"},\"data\":" + output + "},";
            //            break;
            //        case "A Grade Goals":
            //            jsonToReturn += "{\"name\":\"A Grade Goals\",\"headers\":{\"c1\":\"Name\",\"c2\":\"Games\"},\"data\":" + output + "},";
            //            break;
            //        case "Senior Goals":
            //            jsonToReturn += "{\"name\":\"Senior Goals\",\"headers\":{\"c1\":\"Name\",\"c2\":\"Games\"},\"data\":" + output + "},";
            //            break;
            //        case "Junior Goals":
            //            jsonToReturn += "{\"name\":\"Junior Goals\",\"headers\":{\"c1\":\"Name\",\"c2\":\"Games\"},\"data\":" + output + "},";
            //            break;
            //        case "Club Goals":
            //            jsonToReturn += "{\"name\":\"Club Goals\",\"headers\":{\"c1\":\"Name\",\"c2\":\"Games\"},\"data\":" + output + "}"; //Need to update this with comma or no comma when adding more SQL Queries
            //            break;
            //        default:
            //            Console.WriteLine("ERROR: Unknown Milestone Type found - '" + item.Key + "'");
            //            break;
            //    }
            //}
            //jsonToReturn += "]}";
            ////TODO: maybe after all the json has been written, can loop through and if data is empty then remove the whole node?
            //using (StreamWriter sw = new StreamWriter(ConfigurationManager.AppSettings["JsonFilePath"] + fileName))
            //{
            //    sw.Write(jsonToReturn);
            //}
        }
        private static void RetrievePlayers()
        {
            DataTable dt = QueryAccessDatabase("SELECT DISTINCT[Games].Id AS id, FirstName & ' ' & LastName AS name, LastName, FirstName FROM [FHFC Membership List] INNER JOIN [Games] ON Games.Id = [FHFC Membership List].Id WHERE[Games].Id <> 1001 ORDER BY LastName, FirstName, [Games].Id");
            if (dt.Columns.Contains("FirstName"))
            {
                dt.Columns.Remove("FirstName");
            }
            if (dt.Columns.Contains("LastName"))
            {
                dt.Columns.Remove("LastName");
            }
            SaveJsonToFile("{\"players\":" + Newtonsoft.Json.JsonConvert.SerializeObject(dt) + "}", "players.json");
        }
        private static void RetrievePlayerIndividualInfo()
        {
            foreach (DataRow player in QueryAccessDatabase("SELECT DISTINCT Id FROM [Games] WHERE Id <> 1001").Rows)
            {
                string playerId = player["Id"].ToString();
                DataTable dt = QueryAccessDatabase(@"SELECT firstName, middleName, lastName, " +
                    "CINT((SELECT SUM(A) FROM [Games] WHERE ID = " + player["Id"] + ")) AS agrade, " +
                    "CINT((SELECT SUM(A) + SUM(B) + SUM(C) + SUM(OpenW) + SUM(Unknown_Senior) FROM [Games] WHERE ID = " + player["Id"] + ")) AS senior, " +
                    "CINT((SELECT SUM([18]) + SUM([17_5]) + SUM([17]) + SUM([16]) + SUM([16_5sun]) + SUM([16Girls]) + SUM([16sun]) + SUM([15]) + SUM([15sun]) + SUM([14]) + SUM([14Girls]) + SUM([14sun]) + SUM([13]) + SUM(Unknown_Junior) FROM [Games] WHERE ID = " + player["Id"] + ")) AS junior, " +
                    "(SELECT MIN(Year) FROM [Games] WHERE ID = " + player["Id"] + ") AS minYear, " +
                    "(SELECT MAX(Year) FROM [Games] WHERE ID = " + player["Id"] + ") AS maxYear, " +
                    "(SELECT COUNT(Year) FROM [Games] WHERE ID = " + player["Id"] + ") AS seasons, " +
                    "CINT(IIf(IsNull((SELECT SUM(Goals) FROM[Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'A')), 0, (SELECT SUM(Goals) FROM[Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'A'))) AS aGradeGoals, " +
                    "CINT(IIf(IsNull((SELECT SUM(Goals) FROM[Games Details Per Round] WHERE ID = " + player["Id"] + " AND (Grade = 'A' OR Grade = 'B' OR Grade = 'C' OR Grade = 'OpenW'))), 0, (SELECT SUM(Goals) FROM[Games Details Per Round] WHERE ID = " + player["Id"] + " AND (Grade = 'A' OR Grade = 'B' OR Grade = 'C' OR Grade = 'OpenW')))) AS seniorGoals, " +
                    "CINT(IIf(IsNull((SELECT SUM(Goals) FROM[Games Details Per Round] WHERE ID = " + player["Id"] + " AND (Grade = 'U18' OR Grade = 'U17.5' OR Grade = 'U16sun' OR Grade = 'U16Girls' OR Grade = 'U16.5sun' OR Grade = 'U16' OR Grade = 'U15sun' OR Grade = 'U14sun' OR Grade = 'U14Girls' OR Grade = 'U14' OR Grade = 'U13'))), 0, (SELECT SUM(Goals) FROM[Games Details Per Round] WHERE ID = " + player["Id"] + " AND (Grade = 'U18' OR Grade = 'U17.5' OR Grade = 'U16sun' OR Grade = 'U16Girls' OR Grade = 'U16.5sun' OR Grade = 'U16' OR Grade = 'U15sun' OR Grade = 'U14sun' OR Grade = 'U14Girls' OR Grade = 'U14' OR Grade = 'U13')))) AS juniorGoals " +
                    "FROM [FHFC Membership List] " +
                    "WHERE ID = " + player["Id"]);
                string json = Newtonsoft.Json.JsonConvert.SerializeObject(dt);
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
                    "[16Girls] AS u16girlsGames, (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U16Girls' AND [Games Details Per Round].Year = [Games].Year) AS u16girlsGoals, " +
                    "[14Girls] AS u14girlsGames, (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U14Girls' AND [Games Details Per Round].Year = [Games].Year) AS u14girlsGoals " +
                    "FROM [Games] WHERE ID = " + player["Id"] + " " +
                    "UNION " +
                    "SELECT 'Totals', (SELECT SUM(A) FROM [Games] WHERE ID = " + player["Id"] + "), (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'A'), " +
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
                    "(SELECT SUM([16Girls]) FROM [Games] WHERE ID = " + player["Id"] + "), (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U16Girls'), " +
                    "(SELECT SUM([14Girls]) FROM [Games] WHERE ID = " + player["Id"] + "), (SELECT SUM(Goals) FROM [Games Details Per Round] WHERE ID = " + player["Id"] + " AND Grade = 'U14Girls') " +
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
                playerJson = playerJson.Substring(0, playerJson.Length - 1) + "]}";
                SaveJsonToFile(json + playerJson, playerId + ".json");
            }
        }
        private static void SaveJsonToFile(string json, string fileName)
        {
            using (StreamWriter sw = new StreamWriter(ConfigurationManager.AppSettings["JsonFilePath"] + fileName))
            {
                sw.Write(json);
            }
        }
        private static DataTable QueryAccessDatabase(string sql)
        {
            try
            {
                _accessConnection = new OleDbConnection(ConfigurationManager.AppSettings["AccessDatabasePath"]);
            }
            catch (Exception ex)
            {
                Console.Write("Error: Failed to create a database connection. \n{0}", ex.Message);
            }
            DataSet ds = new DataSet();
            try
            {
                OleDbCommand accessCommand = new OleDbCommand(sql, _accessConnection);
                OleDbDataAdapter dataAdapter = new OleDbDataAdapter(accessCommand);
                _accessConnection.Open();
                dataAdapter.Fill(ds);
            }
            catch (Exception ex)
            {
                Console.Write("Error: Failed to retrieve the required data from the database.\n{0}", ex.Message);
                return null;
            }
            finally
            {
                _accessConnection.Close();
            }
            return ds.Tables[0];
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