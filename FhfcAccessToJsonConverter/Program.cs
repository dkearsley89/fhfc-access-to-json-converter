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
            {"Most Senior Games", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT((SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior))) AS a FROM Games INNER JOIN [FHFC Membership List] ON Games.ID = [FHFC Membership List].ID GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior) > 0 ORDER BY (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Games" } },
            {"Most Club Games", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT((SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14Girls])+SUM(Games.[14sun])+SUM(Games.[13])+SUM(Games.unknown_Senior)+SUM(Games.unknown_Junior))) AS a FROM Games INNER JOIN [FHFC Membership List] ON Games.ID = [FHFC Membership List].ID GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName ORDER BY (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14Girls])+SUM(Games.[14sun])+SUM(Games.[13])+SUM(Games.unknown_Senior)+SUM(Games.unknown_Junior)) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Games" } },
            {"Most A Grade Games", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Games.a)) AS a FROM Games INNER JOIN[FHFC Membership List] ON Games.ID = [FHFC Membership List].ID GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Games.a) > 0 ORDER BY SUM(Games.a) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Games" } },

            {"Most B Grade Games", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Games.b)) AS a FROM Games INNER JOIN[FHFC Membership List] ON Games.ID = [FHFC Membership List].ID GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Games.b) > 0 ORDER BY SUM(Games.b) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Games" } },
            {"Most C Grade Games", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Games.c)) AS a FROM Games INNER JOIN[FHFC Membership List] ON Games.ID = [FHFC Membership List].ID GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Games.c) > 0 ORDER BY SUM(Games.c) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Games" } },
            {"Most Open Women's Games", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Games.openW)) AS a FROM Games INNER JOIN[FHFC Membership List] ON Games.ID = [FHFC Membership List].ID GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Games.openW) > 0 ORDER BY SUM(Games.openW) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Games" } },

            {"Most Senior Goals", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'A' OR [Games Details Per Round].Grade = 'B' OR [Games Details Per Round].Grade = 'C' OR [Games Details Per Round].Grade = 'OpenW' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most Club Goals", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most A Grade Goals", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'A' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },

            {"Most B Grade Goals", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'B' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most C Grade Goals", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'C' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most Open Women's Goals", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'OpenW' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },

            {"Avg Goals Per Game - Senior", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, FORMAT(AVG(Goals),\"0.,00\") AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'A' OR [Games Details Per Round].Grade = 'B' OR [Games Details Per Round].Grade = 'C' OR [Games Details Per Round].Grade = 'OpenW' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY AVG(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Avg Goals Per Game - Club", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, FORMAT(AVG(Goals),\"0.,00\") AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY AVG(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Avg Goals Per Game - A Grade", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, FORMAT(AVG(Goals),\"0.,00\") AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'A' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING AVG(Goals) > 0 ORDER BY AVG(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },

            {"Avg Goals Per Game - B Grade", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, FORMAT(AVG(Goals),\"0.,00\") AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'B' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY AVG(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Avg Goals Per Game - C Grade", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, FORMAT(AVG(Goals),\"0.,00\") AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'C' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY AVG(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Avg Goals Per Game - Open Women's", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, FORMAT(AVG(Goals),\"0.,00\") AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'OpenW' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY AVG(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },

            {"Most Goals in a Season - Senior", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'A' OR [Games Details Per Round].Grade = 'B' OR [Games Details Per Round].Grade = 'C' OR [Games Details Per Round].Grade = 'OpenW' GROUP BY [FHFC Membership List].id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most Goals in a Season - Club", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID GROUP BY [FHFC Membership List].id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most Goals in a Season - A Grade", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'A' GROUP BY [FHFC Membership List].id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },

            {"Most Goals in a Season - B Grade", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'B' GROUP BY [FHFC Membership List].id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most Goals in a Season - C Grade", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'C' GROUP BY [FHFC Membership List].id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most Goals in a Season - Open Women's", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'OpenW' GROUP BY [FHFC Membership List].id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },

            {"Most Goals in a Game - Senior", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, Year, Round, Grade, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, Goals AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'A' OR [Games Details Per Round].Grade = 'B' OR [Games Details Per Round].Grade = 'C' OR [Games Details Per Round].Grade = 'OpenW' GROUP BY [FHFC Membership List].id, Year, Round, Grade, Goals, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING Goals > 0 ORDER BY Goals DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most Goals in a Game - Club", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, Year, Round, Grade, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, Goals AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID GROUP BY [FHFC Membership List].id, Year, Round, Grade, Goals, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING Goals > 0 ORDER BY Goals DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most Goals in a Game - A Grade", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, Year, Round, Grade, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, Goals AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'A' GROUP BY [FHFC Membership List].id, Year, Round, Grade, Goals, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING Goals > 0 ORDER BY Goals DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },

            {"Most Goals in a Game - B Grade", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, Year, Round, Grade, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, Goals AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'B' GROUP BY [FHFC Membership List].id, Year, Round, Grade, Goals, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING Goals > 0 ORDER BY Goals DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most Goals in a Game - C Grade", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, Year, Round, Grade, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, Goals AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'C' GROUP BY [FHFC Membership List].id, Year, Round, Grade, Goals, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING Goals > 0 ORDER BY Goals DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most Goals in a Game - Open Women's", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, Year, Round, Grade, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, Goals AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'OpenW' GROUP BY [FHFC Membership List].id, Year, Round, Grade, Goals, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING Goals > 0 ORDER BY Goals DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },

            {"Most Junior Games", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT((SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14Girls])+SUM(Games.[14sun])+SUM(Games.[13])+SUM(Games.unknown_Junior))) AS a FROM Games INNER JOIN [FHFC Membership List] ON Games.ID = [FHFC Membership List].ID GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14Girls])+SUM(Games.[14sun])+SUM(Games.[13])+SUM(Games.unknown_Junior) > 0 ORDER BY (SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14Girls])+SUM(Games.[14sun])+SUM(Games.[13])+SUM(Games.unknown_Junior)) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Games" } },
            {"Most Junior Goals", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'U18' OR [Games Details Per Round].Grade = 'U17.5' OR [Games Details Per Round].Grade = 'U16' OR [Games Details Per Round].Grade = 'U16sun' OR [Games Details Per Round].Grade = 'U15' OR [Games Details Per Round].Grade = 'U15sun' OR [Games Details Per Round].Grade = 'U14' OR [Games Details Per Round].Grade = 'U14sun' OR [Games Details Per Round].Grade = 'U14Girls' OR [Games Details Per Round].Grade = 'U13' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Avg Goals Per Game - Junior", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, FORMAT(AVG(Goals),\"0.,00\") AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'U18' OR [Games Details Per Round].Grade = 'U17.5' OR [Games Details Per Round].Grade = 'U16' OR [Games Details Per Round].Grade = 'U16sun' OR [Games Details Per Round].Grade = 'U15' OR [Games Details Per Round].Grade = 'U15sun' OR [Games Details Per Round].Grade = 'U14' OR [Games Details Per Round].Grade = 'U14sun' OR [Games Details Per Round].Grade = 'U14Girls' OR [Games Details Per Round].Grade = 'U13' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY AVG(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },

            {"Most Goals in a Season - Junior", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, CINT(SUM(Goals)) AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'U18' OR [Games Details Per Round].Grade = 'U17.5' OR [Games Details Per Round].Grade = 'U16' OR [Games Details Per Round].Grade = 'U16sun' OR [Games Details Per Round].Grade = 'U15' OR [Games Details Per Round].Grade = 'U15sun' OR [Games Details Per Round].Grade = 'U14' OR [Games Details Per Round].Grade = 'U14sun' OR [Games Details Per Round].Grade = 'U14Girls' OR [Games Details Per Round].Grade = 'U13' GROUP BY [FHFC Membership List].id, Year, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING SUM(Goals) > 0 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most Goals in a Game - Junior", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, Year, Round, Grade, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, Goals AS a FROM [Games Details Per Round] INNER JOIN [FHFC Membership List] ON [Games Details Per Round].ID = [FHFC Membership List].ID WHERE [Games Details Per Round].Grade = 'U18' OR [Games Details Per Round].Grade = 'U17.5' OR [Games Details Per Round].Grade = 'U16' OR [Games Details Per Round].Grade = 'U16sun' OR [Games Details Per Round].Grade = 'U15' OR [Games Details Per Round].Grade = 'U15sun' OR [Games Details Per Round].Grade = 'U14' OR [Games Details Per Round].Grade = 'U14sun' OR [Games Details Per Round].Grade = 'U14Girls' OR [Games Details Per Round].Grade = 'U13' GROUP BY [FHFC Membership List].id, Year, Round, Grade, Goals, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING Goals > 0 ORDER BY Goals DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Goals" } },
            {"Most Seasons Played", new [] { "SELECT TOP [NumberOfRecords] [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, COUNT(Games.Id) AS a FROM Games INNER JOIN[FHFC Membership List] ON Games.ID = [FHFC Membership List].ID GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName ORDER BY COUNT(Games.Id) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName", "Seasons" } }
        };
        //private readonly static Dictionary<string, string> _milestoneSqlStatements = new Dictionary<string, string>()
        //{
        //    {"A Grade Games", "SELECT [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, SUM(Games.a) AS a FROM Games INNER JOIN [FHFC Membership List] ON Games.ID = [FHFC Membership List].ID WHERE [FHFC Membership List].currentlyActive = true GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING (SUM(Games.a) BETWEEN 45 AND 49 OR SUM(Games.a) BETWEEN 90 AND 99 OR SUM(Games.a) BETWEEN 145 AND 149 OR SUM(Games.a) BETWEEN 190 AND 199 OR SUM(Games.a) BETWEEN 245 AND 249 OR SUM(Games.a) BETWEEN 290 AND 299) AND MAX(Games.Year) > Year(Now())-2 ORDER BY SUM(Games.a) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName"},
        //    {"B Grade Games", "SELECT [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, SUM(Games.b) AS a FROM Games INNER JOIN [FHFC Membership List] ON Games.ID = [FHFC Membership List].ID WHERE [FHFC Membership List].currentlyActive = true GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING (SUM(Games.b) BETWEEN 45 AND 49 OR SUM(Games.b) BETWEEN 90 AND 99 OR SUM(Games.b) BETWEEN 145 AND 149 OR SUM(Games.b) BETWEEN 190 AND 199 OR SUM(Games.b) BETWEEN 245 AND 249 OR SUM(Games.b) BETWEEN 290 AND 299) AND MAX(Games.Year) > Year(Now())-2 ORDER BY SUM(Games.b) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName"},
        //    {"C Grade Games", "SELECT [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, SUM(Games.c) AS a FROM Games INNER JOIN [FHFC Membership List] ON Games.ID = [FHFC Membership List].ID WHERE [FHFC Membership List].currentlyActive = true GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING (SUM(Games.c) BETWEEN 45 AND 49 OR SUM(Games.c) BETWEEN 90 AND 99 OR SUM(Games.c) BETWEEN 145 AND 149 OR SUM(Games.c) BETWEEN 190 AND 199 OR SUM(Games.c) BETWEEN 245 AND 249 OR SUM(Games.c) BETWEEN 290 AND 299) AND MAX(Games.Year) > Year(Now())-2 ORDER BY SUM(Games.c) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName"},
        //    {"Open Women Games", "SELECT [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, SUM(Games.openW) AS a FROM Games INNER JOIN [FHFC Membership List] ON Games.ID = [FHFC Membership List].ID WHERE [FHFC Membership List].currentlyActive = true GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING (SUM(Games.openW) BETWEEN 45 AND 49 OR SUM(Games.openW) BETWEEN 90 AND 99 OR SUM(Games.openW) BETWEEN 145 AND 149 OR SUM(Games.openW) BETWEEN 190 AND 199 OR SUM(Games.openW) BETWEEN 245 AND 249 OR SUM(Games.openW) BETWEEN 290 AND 299) AND MAX(Games.Year) > Year(Now())-2 ORDER BY SUM(Games.openW) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName"},
        //    {"Senior Games", "SELECT [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)) AS a FROM Games INNER JOIN [FHFC Membership List] ON Games.ID = [FHFC Membership List].ID WHERE [FHFC Membership List].currentlyActive = true GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING ((SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)) BETWEEN 45 AND 49 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)) BETWEEN 90 AND 99 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)) BETWEEN 145 AND 149 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)) BETWEEN 190 AND 199 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)) BETWEEN 245 AND 249 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)) BETWEEN 290 AND 299 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)) BETWEEN 345 AND 349OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)) BETWEEN 390 AND 399OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)) BETWEEN 445 AND 449 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)) BETWEEN 490 AND 499) AND MAX(Games.Year) > Year(Now())-2 ORDER BY (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName"},
        //    {"Junior Games", "SELECT [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, (SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) AS a FROM Games INNER JOIN [FHFC Membership List] ON Games.ID = [FHFC Membership List].ID WHERE [FHFC Membership List].currentlyActive = true AND [FHFC Membership List].membershipStatus = 'jp' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING ((SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 45 AND 49 OR (SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 90 AND 99 OR (SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 145 AND 149 OR (SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 190 AND 199) AND MAX(Games.Year) > Year(Now())-2 ORDER BY (SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName"},
        //    {"Club Games", "SELECT [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) AS a FROM Games INNER JOIN [FHFC Membership List] ON Games.ID = [FHFC Membership List].ID WHERE [FHFC Membership List].currentlyActive = true GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING ((SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 45 AND 49 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 90 AND 99 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 145 AND 149 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 190 AND 199 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 245 AND 249 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 290 AND 299 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 345 AND 349OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 390 AND 399OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 445 AND 449 OR (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) BETWEEN 490 AND 499) AND MAX(Games.Year) > Year(Now())-2 ORDER BY (SUM(Games.a)+SUM(Games.b)+SUM(Games.c)+SUM(Games.openW)+SUM(Games.unknown_Senior)+SUM(Games.[18])+SUM(Games.[17_5])+SUM(Games.[17])+SUM(Games.[16])+SUM(Games.[16sun])+SUM(Games.[15])+SUM(Games.[15sun])+SUM(Games.[14])+SUM(Games.[14sun])+SUM(Games.[14girls])+SUM(Games.[13])+SUM(Games.[unknown_Junior])) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName"},
        //    {"A Grade Goals", "SELECT [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, SUM(Goals) AS a FROM [Games Details per round] INNER JOIN [FHFC Membership List] ON [Games Details per round].ID = [FHFC Membership List].ID WHERE [FHFC Membership List].currentlyActive = true AND [Games Details per round].Grade = 'A' GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING (SUM(Goals) BETWEEN 80 AND 99 OR  SUM(Goals) BETWEEN 180 AND 199 OR SUM(Goals) BETWEEN 280 AND 299 OR SUM(Goals) BETWEEN 380 AND 399 OR SUM(Goals) BETWEEN 480 AND 499 OR SUM(Goals) BETWEEN 580 AND 599 OR SUM(Goals) BETWEEN 680 AND 699 OR SUM(Goals) BETWEEN 780 AND 799 OR SUM(Goals) BETWEEN 880 AND 899 OR SUM(Goals) BETWEEN 980 AND 999) AND MAX(Year) > Year(Now())-2 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName"},
        //    {"Senior Goals", "SELECT [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, SUM(Goals) AS a FROM [Games Details per round] INNER JOIN [FHFC Membership List] ON [Games Details per round].ID = [FHFC Membership List].ID WHERE [FHFC Membership List].currentlyActive = true AND ([Games Details per round].Grade = 'A' OR [Games Details per round].Grade = 'B' OR [Games Details per round].Grade = 'C' OR [Games Details per round].Grade = 'OpenW') GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING (SUM(Goals) BETWEEN 80 AND 99 OR  SUM(Goals) BETWEEN 180 AND 199 OR SUM(Goals) BETWEEN 280 AND 299 OR SUM(Goals) BETWEEN 380 AND 399 OR SUM(Goals) BETWEEN 480 AND 499 OR SUM(Goals) BETWEEN 580 AND 599 OR SUM(Goals) BETWEEN 680 AND 699 OR SUM(Goals) BETWEEN 780 AND 799 OR SUM(Goals) BETWEEN 880 AND 899 OR SUM(Goals) BETWEEN 980 AND 999) AND MAX(Year) > Year(Now())-2 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName"},
        //    {"Junior Goals", "SELECT [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, SUM(Goals) AS a FROM [Games Details per round] INNER JOIN [FHFC Membership List] ON [Games Details per round].ID = [FHFC Membership List].ID WHERE [FHFC Membership List].currentlyActive = true AND [FHFC Membership List].membershipStatus = 'jp' AND ([Games Details per round].Grade = 'U18' OR [Games Details per round].Grade = 'U17.5' OR [Games Details per round].Grade = 'U17' OR [Games Details per round].Grade = 'U16' OR [Games Details per round].Grade = 'U16sun' OR [Games Details per round].Grade = 'U15' OR [Games Details per round].Grade = 'U15sun' OR [Games Details per round].Grade = 'U14' OR [Games Details per round].Grade = 'U14sun' OR [Games Details per round].Grade = 'U14Girls' OR [Games Details per round].Grade = 'U13') GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING (SUM(Goals) BETWEEN 80 AND 99 OR  SUM(Goals) BETWEEN 180 AND 199 OR SUM(Goals) BETWEEN 280 AND 299 OR SUM(Goals) BETWEEN 380 AND 399 OR SUM(Goals) BETWEEN 480 AND 499 OR SUM(Goals) BETWEEN 580 AND 599 OR SUM(Goals) BETWEEN 680 AND 699 OR SUM(Goals) BETWEEN 780 AND 799 OR SUM(Goals) BETWEEN 880 AND 899 OR SUM(Goals) BETWEEN 980 AND 999) AND MAX(Year) > Year(Now())-2 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName"},
        //    {"Club Goals", "SELECT [FHFC Membership List].id As Id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName AS n, SUM(Goals) AS a FROM [Games Details per round] INNER JOIN [FHFC Membership List] ON [Games Details per round].ID = [FHFC Membership List].ID WHERE [FHFC Membership List].currentlyActive = true GROUP BY [FHFC Membership List].id, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName HAVING (SUM(Goals) BETWEEN 80 AND 99 OR  SUM(Goals) BETWEEN 180 AND 199 OR SUM(Goals) BETWEEN 280 AND 299 OR SUM(Goals) BETWEEN 380 AND 399 OR SUM(Goals) BETWEEN 480 AND 499 OR SUM(Goals) BETWEEN 580 AND 599 OR SUM(Goals) BETWEEN 680 AND 699 OR SUM(Goals) BETWEEN 780 AND 799 OR SUM(Goals) BETWEEN 880 AND 899 OR SUM(Goals) BETWEEN 980 AND 999) AND MAX(Year) > Year(Now())-2 ORDER BY SUM(Goals) DESC, [FHFC Membership List].firstName & \" \" & [FHFC Membership List].lastName"}
        //};
        private static OleDbConnection _accessConnection;
        private static void Main()
        {
            Console.WriteLine("Starting FhfcAccessToJsonConverter");
            RetrieveHomePageInfo();
            RetrieveSummaryRecords();
            RetrieveRecords();
            //RetrieveMilestones("milestones.json");
            //RetrieveMembers("members.json");
            Console.WriteLine("Finishing FhfcAccessToJsonConverter");
            Console.Read();
        }
        private static void RetrieveHomePageInfo()
        {
            string[] homePageRecordsToDisplay = new[] { "Most Senior Games", "Most A Grade Games", "Most Club Games", "Most Junior Games", "Most Goals in a Season - Club", "Most Club Goals" };
            string jsonToReturn = "{\"lastUpdated\": \"" + String.Format(DateTime.Now.ToString("MMMM dd{0} yyyy"), GetDaySuffix(DateTime.Now.Day)) + "\",\"records\":[";
            foreach (KeyValuePair<string, string[]> item in RecordSqlStatements)
            {
                if (homePageRecordsToDisplay.Contains(item.Key))
                {
                    DataTable dt = QueryAccessDatabase(item.Value[0].Replace("[NumberOfRecords]", "1"));
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
                    jsonToReturn += "{\"name\":\"" + item.Key + "\",\"label\":\"" + item.Value[1] + "\",\"data\":" + Newtonsoft.Json.JsonConvert.SerializeObject(dt) + "},";
                }
            }
            jsonToReturn = jsonToReturn.TrimEnd(',') + "]}";
            SaveJsonToFile(jsonToReturn, "homePageInfo.json");
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
        private static void RetrieveMembers(string fileName)
        {
            //SaveJsonToFile("{\"members\":" + Newtonsoft.Json.JsonConvert.SerializeObject(QueryAccessDatabase("SELECT DISTINCT [Games].Id, FirstName As f, LastName As l FROM [FHFC Membership List] INNER JOIN [Games] ON Games.Id = [FHFC Membership List].Id WHERE [Games].Id <> 1001 ORDER BY LastName ASC, FirstName ASC, [Games].Id ASC")) + "}", fileName);
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