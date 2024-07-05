using System.Data;
using Microsoft.Data.SqlClient;
using Microsoft.IdentityModel.Tokens;
using WordProcessor.Table1.Entities;

namespace WordProcessor.Table1;

public static class Connection
{
    public static string conString =  @"Data Source=.\SQLEXPRESS; Initial Catalog=FilesDatabase; Integrated Security=true; TrustServerCertificate=True;";

    public static List<TrainedStudent> GetParticipantsForContract(string contractNumber)
    {
        List<TrainedStudent> participants = new List<TrainedStudent>();

        string cmdString = "SELECT * FROM Table_1 WHERE VUZINN = (SELECT INN FROM Accellerations WHERE ContractID = '" + contractNumber + "')";

        SqlConnection con = new SqlConnection(conString);

        SqlCommand cmd = new SqlCommand(cmdString, con);

        cmd.CommandType = CommandType.Text;
        
        con.Open();

        SqlDataReader reader = cmd.ExecuteReader();

        var studentID = 1;
        
        while (reader.Read())
        {
            TrainedStudent p = new TrainedStudent();
            p.Number = studentID;
            p.LeaderId = Convert.ToInt64(reader[0]);
            p.FIO = string.IsNullOrEmpty(reader[1].ToString()) ? "-" : reader[1].ToString() ?? "-";
            p.EventsId = string.IsNullOrEmpty(reader[2].ToString()) ? "-" : reader[2].ToString() ?? "-";
            p.Count = Convert.ToInt32(reader[3]);
            p.StartUp = string.IsNullOrEmpty(reader[4].ToString()) ? "-" : reader[4].ToString() ?? "-";
            p.Link = string.IsNullOrEmpty(reader[5].ToString()) ? "-" : reader[5].ToString() ?? "-";
            studentID++;
            participants.Add(p);
        }

        con.Close();

        return participants;
    }
    public static List<Event> GetEventsForContract(string contractNumber)
    {
        List<Event> events = new List<Event>();

        string cmdString = "SELECT * FROM Table_2 WHERE VUZEventProviderINN = (SELECT INN FROM Accellerations WHERE ContractID = '" + contractNumber + "') ORDER BY StartDate ASC";

        SqlConnection con = new SqlConnection(conString);

        SqlCommand cmd = new SqlCommand(cmdString, con);

        cmd.CommandType = CommandType.Text;
        
        con.Open();

        SqlDataReader reader = cmd.ExecuteReader();

        var eventID = 1;
        
        while (reader.Read())
        {
            Event e = new Event();
            e.Number = eventID;
            e.Name = string.IsNullOrEmpty(reader[0].ToString()) ? "-" : reader[0].ToString() ?? "-";
            e.LeaderId = string.IsNullOrEmpty(reader[1].ToString()) ? "-" : reader[1].ToString() ?? "-";
            e.Link = string.IsNullOrEmpty(reader[2].ToString()) ? "-" : reader[2].ToString() ?? "-";
            e.DateStart = string.IsNullOrEmpty(reader[3].ToString()) ? "-" : Convert.ToDateTime(reader[3]).ToString("dd.MM.yyyy HH:mm");
            e.Format = string.IsNullOrEmpty(reader[4].ToString()) ? "-" : reader[4].ToString() ?? "-";
            e.CountOfParticipants = string.IsNullOrEmpty(reader[5].ToString()) ? 0 : Convert.ToInt64(reader[5]);
            e.LeaderIdNumber = string.IsNullOrEmpty(reader[6].ToString()) ? "-" : reader[6].ToString() ?? "-";
            eventID++;
            events.Add(e);
        }

        con.Close();

        return events;
    }
    
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
            s.Number = startupID;
            var startupIDForProcedure = Convert.ToInt32(reader[0]);
            s.Name = string.IsNullOrEmpty(reader[1].ToString()) ? "-" : reader[1].ToString() ?? "-";
            s.Link = string.IsNullOrEmpty(reader[2].ToString()) ? "-" : reader[2].ToString() ?? "-";
            s.HasSign = "";
            s.Category = "";
            s.Participants = GetParticipantsForStartup(startupIDForProcedure);
            startupID++;
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
            p.Number = studentID;
            p.LeaderId = Convert.ToInt64(reader[0]);
            p.FIO = string.IsNullOrEmpty(reader[1].ToString()) ? "-" : reader[1].ToString() ?? "-";
            p.EventsId = string.IsNullOrEmpty(reader[2].ToString()) ? "-" : reader[2].ToString() ?? "-";
            p.Count = Convert.ToInt32(reader[3]);
            p.StartUp = string.IsNullOrEmpty(reader[4].ToString()) ? "-" : reader[4].ToString() ?? "-";
            p.Link = string.IsNullOrEmpty(reader[5].ToString()) ? "-" : reader[5].ToString() ?? "-";
            studentID++;
            participants.Add(p);
        }

        con.Close();

        return participants;
    }
    
    public static List<ErrorTable1> GetErrors1ForContract(string contractID)
    {
        List<ErrorTable1> participants = new List<ErrorTable1>();

        string cmdString = "SELECT * FROM ErrorTable1 WHERE ContractID = '" + contractID + "'";

        SqlConnection con = new SqlConnection(conString);

        SqlCommand cmd = new SqlCommand(cmdString, con);

        cmd.CommandType = CommandType.Text;
        
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
            p.Documents = "Необходимо запросить документы, подтверждающие участие студента в акселерационной программе либо внести соответствующие корректировки в отчетные документы";
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
            e.DateStart = string.IsNullOrEmpty(reader[3].ToString()) ? "-" : Convert.ToDateTime(reader[3]).ToString("dd.MM.yyyy HH:mm");
            e.Format = string.IsNullOrEmpty(reader[4].ToString()) ? "-" : reader[4].ToString() ?? "-";
            e.CountOfParticipants = string.IsNullOrEmpty(reader[5].ToString()) ? 0 : Convert.ToInt64(reader[5]);
            e.LeaderIdNumber = string.IsNullOrEmpty(reader[6].ToString()) ? "-" : reader[6].ToString() ?? "-";
            events.Add(e);
        }

        con.Close();

        return events.OrderBy(e =>
            {
                DateTime parsedDate;
                return DateTime.TryParse(e.DateStart, out parsedDate) ? parsedDate : DateTime.MinValue;
            })
            .Select((e, index) => { e.Number = index + 1; return e; })
            .ToList();
    }
}