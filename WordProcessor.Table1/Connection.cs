using System.Data;
using Microsoft.Data.SqlClient;
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
            p.FIO = reader[1].ToString();
            p.EventsId = reader[2].ToString();
            p.Count = Convert.ToInt32(reader[3]);
            p.StartUp = reader[4].ToString();
            p.Link = reader[5].ToString();
            studentID++;
            participants.Add(p);
        }

        con.Close();

        return participants;
    }
    
    
    public static List<Event> GetEventsForContract(string contractNumber)
    {
        List<Event> events = new List<Event>();

        string cmdString = "SELECT * FROM Table_2 WHERE VUZEventProviderINN = (SELECT INN FROM Accellerations WHERE ContractID = '" + contractNumber + "')";

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
            e.Name = reader[0].ToString();
            e.LeaderId = reader[1].ToString();
            e.Link = reader[2].ToString();
            e.DateStart = Convert.ToDateTime(reader[3]);
            e.Format = reader[4].ToString();
            e.CountOfParticipants = Convert.ToInt64(reader[5]);
            e.LeaderIdNumber = reader[6].ToString();
            eventID++;
            events.Add(e);
        }

        con.Close();

        return events;
    }
}