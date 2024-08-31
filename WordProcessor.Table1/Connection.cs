using System.Data;
using DocumentFormat.OpenXml.Bibliography;
using Microsoft.Data.SqlClient;
using WordProcessor.Table1.Entities;

namespace WordProcessor.Table1;

public static class Connection
{
    public static string conString =
        @"Data Source=.\SQLEXPRESS; Initial Catalog=FilesDatabase; Integrated Security=true; TrustServerCertificate=True;";

    public static List<Startup> GetStartupsForContract(string contractNumber)
    {
        List<Startup> startups = new List<Startup>();

        string cmdString = "GetStartupsForContract";

        SqlConnection con = new SqlConnection(conString);

        SqlCommand cmd = new SqlCommand(cmdString, con);

        cmd.CommandType = CommandType.StoredProcedure;

        SqlParameter idParam = new SqlParameter
        {
            ParameterName = "@id",
            Value = contractNumber
        };
        cmd.Parameters.Add(idParam);

        con.Open();

        SqlDataReader reader = cmd.ExecuteReader();

        var startupID = 1;

        while (reader.Read())
        {
            Startup s = new Startup();
            s.Number = startupID.ToString();
            s.Name = string.IsNullOrEmpty(reader[0].ToString()) ? "-" : reader[0].ToString() ?? "-";
            s.Link = string.IsNullOrEmpty(reader[1].ToString()) ? "-" : reader[1].ToString() ?? "-";
            s.DupeCount = reader[2].ToString() ?? 0.ToString();
            s.HasSign = "да";
            s.Category = "2";
            
            var startupIDForProcedure = GetStartupByLink(s.Link, contractNumber);
            
            startupID++;
            
            s.Participants = GetParticipantsForStartup(startupIDForProcedure);
            startups.Add(s);
        }

        con.Close();

        return startups;
    }

    public static List<Participant> GetParticipantsForStartup(int startupID)
    {
        List<Participant> participants = new List<Participant>();

        string cmdString = "GetParticipantsForStartup";

        SqlConnection con = new SqlConnection(conString);

        SqlCommand cmd = new SqlCommand(cmdString, con);

        cmd.CommandType = CommandType.StoredProcedure;

        SqlParameter idParam = new SqlParameter
        {
            ParameterName = "@id",
            Value = startupID
        };
        cmd.Parameters.Add(idParam);

        con.Open();

        SqlDataReader reader = cmd.ExecuteReader();

        while (reader.Read())
        {
            Participant p = new Participant();
            p.Name = string.IsNullOrEmpty(reader[1].ToString()) ? "" : reader[1].ToString() ?? "";
            p.LeaderID = string.IsNullOrEmpty(reader[2].ToString()) ? "-" : reader[2].ToString() ?? "-";
            p.EventIDs = string.IsNullOrEmpty(reader[3].ToString()) ? "-" : reader[3].ToString() ?? "-";
            participants.Add(p);
        }

        con.Close();

        return participants;
    }

    public static List<TrainedStudent> GetNewTable1ForContract(string contractID)
    {
        List<TrainedStudent> participants = new List<TrainedStudent>();

        string cmdString = "GetNewTable1ForContract";

        SqlConnection con = new SqlConnection(conString);

        SqlCommand cmd = new SqlCommand(cmdString, con);

        cmd.CommandType = CommandType.StoredProcedure;

        cmd.CommandTimeout = 1000000;

        SqlParameter idParam = new SqlParameter
        {
            ParameterName = "@id",
            Value = contractID
        };
        cmd.Parameters.Add(idParam);

        con.Open();

        SqlDataReader reader = cmd.ExecuteReader();

        var studentID = 1;

        while (reader.Read())
        {
            TrainedStudent p = new TrainedStudent();
            p.Number = studentID.ToString();
            p.LeaderId = string.IsNullOrEmpty(reader[0].ToString()) ? "" : reader[0].ToString() ?? "";
            p.FIO = string.IsNullOrEmpty(reader[1].ToString()) ? "-" : reader[1].ToString() ?? "-";
            p.EventsId = string.IsNullOrEmpty(reader[2].ToString()) ? "-" : reader[2].ToString() ?? "-";
            p.Count = string.IsNullOrEmpty(reader[3].ToString()) ? "" : reader[3].ToString() ?? "";
            p.StartUp = string.IsNullOrEmpty(reader[4].ToString()) ? "-" : reader[4].ToString() ?? "-";
            p.Link = string.IsNullOrEmpty(reader[5].ToString()) ? "-" : reader[5].ToString() ?? "-";
            studentID++;
            participants.Add(p);
        }

        con.Close();

        foreach (var p in participants)
        {
            if (p.FIO == '-'.ToString())
            {
                var name = GetStudentByID(p.LeaderId.ToString(), contractID);
                if (name != "#Н/Д")
                {
                    p.FIO = name;
                }
            }
        }
        
        return participants;
    }

    public static List<ErrorTable2> GetErrors2ForContract(string contractID)
    {
        List<ErrorTable2> events = new List<ErrorTable2>();

        string cmdString = "GetEventsWithErrors";

        SqlConnection con = new SqlConnection(conString);

        SqlCommand cmd = new SqlCommand(cmdString, con);

        cmd.CommandType = CommandType.StoredProcedure;

        SqlParameter idParam = new SqlParameter
        {
            ParameterName = "@id",
            Value = contractID
        };
        cmd.Parameters.Add(idParam);

        con.Open();

        SqlDataReader reader = cmd.ExecuteReader();

        var errorID = 1;

        while (reader.Read())
        {
            ErrorTable2 p = new ErrorTable2();
            p.Number = errorID.ToString();
            p.Name = reader[0].ToString();
            p.Link = reader[1].ToString();
            p.Reason = reader[2].ToString();
            p.Documents = "-";
            p.Remark = "-";
            p.Comment = "-";
            errorID++;
            events.Add(p);
        }

        con.Close();

        return events;
    }

    public static List<ErrorTable3> GetErrors3ForContract(string contractID)
    {
        List<ErrorTable3> startups = new List<ErrorTable3>();

        string cmdString = "GetStartupsWithErrors";

        SqlConnection con = new SqlConnection(conString);

        SqlCommand cmd = new SqlCommand(cmdString, con);

        cmd.CommandType = CommandType.StoredProcedure;

        SqlParameter idParam = new SqlParameter
        {
            ParameterName = "@id",
            Value = contractID
        };
        cmd.Parameters.Add(idParam);

        con.Open();

        SqlDataReader reader = cmd.ExecuteReader();

        var errorID = 1;

        while (reader.Read())
        {
            ErrorTable3 p = new ErrorTable3();
            p.Number = errorID.ToString();
            p.Name = reader[1].ToString();
            p.Link = reader[2].ToString();
            p.Reason = reader[3].ToString();
            p.Documents = "-";
            p.Remark = "-";
            p.Comment = "-";
            errorID++;
            startups.Add(p);
        }

        con.Close();

        return startups;
    }

    public static List<ErrorTable1> GetErrorsForContract(string contractID)
    {
        List<ErrorTable1> participants = new List<ErrorTable1>();

        string cmdString = "GetErrorsForContract";

        SqlConnection con = new SqlConnection(conString);

        SqlCommand cmd = new SqlCommand(cmdString, con);

        cmd.CommandType = CommandType.StoredProcedure;

        SqlParameter idParam = new SqlParameter
        {
            ParameterName = "@id",
            Value = contractID
        };
        cmd.Parameters.Add(idParam);

        con.Open();

        SqlDataReader reader = cmd.ExecuteReader();

        var studentID = 1;

        while (reader.Read())
        {
            ErrorTable1 p = new ErrorTable1();
            p.Number = studentID.ToString();
            p.Name = reader[0].ToString();
            p.LeaderLink = reader[1].ToString();
            p.Reason = reader[2].ToString();
            p.Documents =
                "Необходимо запросить документы, подтверждающие участие студента в акселерационной программе либо внести соответствующие корректировки в отчетные документы";
            p.Remark = "";
            p.Comment = "";
            studentID++;
            participants.Add(p);
        }

        con.Close();

        return participants;
    }

    public static List<Event> GetNewTable2ForContract(string contractID)
    {
        List<Event> events = new List<Event>();

        string cmdString = "GetNewTable2ForContract";

        SqlConnection con = new SqlConnection(conString);

        SqlCommand cmd = new SqlCommand(cmdString, con);

        cmd.CommandType = CommandType.StoredProcedure;

        SqlParameter idParam = new SqlParameter
        {
            ParameterName = "@id",
            Value = contractID
        };
        cmd.Parameters.Add(idParam);

        con.Open();

        SqlDataReader reader = cmd.ExecuteReader();

        while (reader.Read())
        {
            Event e = new Event();
            e.Name = string.IsNullOrEmpty(reader[0].ToString()) ? "-" : reader[0].ToString() ?? "-";
            e.LeaderId = string.IsNullOrEmpty(reader[1].ToString()) ? "-" : reader[1].ToString() ?? "-";
            e.Link = string.IsNullOrEmpty(reader[2].ToString()) ? "-" : reader[2].ToString() ?? "-";
            e.DateStart = string.IsNullOrEmpty(reader[3].ToString())
                ? "-"
                : Convert.ToDateTime(reader[3]).ToString("dd.MM.yyyy HH:mm");
            e.Format = string.IsNullOrEmpty(reader[4].ToString()) ? "-" : reader[4].ToString() ?? "-";
            e.CountOfParticipants = string.IsNullOrEmpty(reader[5].ToString()) ? 0.ToString() : reader[5].ToString();
            e.LeaderIdNumber = string.IsNullOrEmpty(reader[6].ToString()) ? "-" : reader[6].ToString() ?? "-";
            events.Add(e);
        }

        con.Close();

        return events.OrderBy(e =>
            {
                DateTime parsedDate;
                return DateTime.TryParse(e.DateStart, out parsedDate) ? parsedDate : DateTime.MinValue;
            })
            .Select((e, index) =>
            {
                e.Number = (index + 1).ToString();
                return e;
            })
            .ToList();
    }

    public static int GetStartupByLink(string link, string contractID)
    {
        string cmdString = "GetStartupByLink";

        SqlConnection con = new SqlConnection(conString);

        SqlCommand cmd = new SqlCommand(cmdString, con);

        cmd.CommandType = CommandType.StoredProcedure;

        SqlParameter idParam = new SqlParameter
        {
            ParameterName = "@id",
            Value = contractID
        };
        cmd.Parameters.Add(idParam);

        SqlParameter linkParam = new SqlParameter
        {
            ParameterName = "@link",
            Value = link
        };
        cmd.Parameters.Add(linkParam);

        con.Open();

       
        var startupID = Convert.ToInt32(cmd.ExecuteScalar());
        
        con.Close();

        return startupID;
    }
    
    public static string GetStudentByID(string lid, string contractID)
    {
        string cmdString = "GetStudentByID";

        SqlConnection con = new SqlConnection(conString);

        SqlCommand cmd = new SqlCommand(cmdString, con);

        cmd.CommandType = CommandType.StoredProcedure;

        SqlParameter idParam = new SqlParameter
        {
            ParameterName = "@id",
            Value = contractID
        };
        cmd.Parameters.Add(idParam);

        SqlParameter lParam = new SqlParameter
        {
            ParameterName = "@lid",
            Value = lid
        };
        cmd.Parameters.Add(lParam);

        con.Open();

       
        var name = cmd.ExecuteScalar() ?? "";
        
        con.Close();

        return name.ToString();
    }
    
    public static string GetEventByID(string eid, string contractID)
    {
        string cmdString = "GetEventByID";

        SqlConnection con = new SqlConnection(conString);

        SqlCommand cmd = new SqlCommand(cmdString, con);

        cmd.CommandType = CommandType.StoredProcedure;

        SqlParameter idParam = new SqlParameter
        {
            ParameterName = "@id",
            Value = contractID
        };
        cmd.Parameters.Add(idParam);

        SqlParameter eParam = new SqlParameter
        {
            ParameterName = "@eid",
            Value = eid
        };
        cmd.Parameters.Add(eParam);

        con.Open();

       
        var name = cmd.ExecuteScalar() ?? "";
        
        con.Close();

        return name.ToString();
    }

    public static bool CheckIfHasGoodParticipants(string contractID, string link)
    {

        var startupID = GetStartupByLink(link, contractID);
        
        var participants = GetParticipantsForStartup(startupID);

        if (participants.Count() == participants.Count(p => p.EventIDs == '-'.ToString()))
            return false;
        else
            return true;
    }
}