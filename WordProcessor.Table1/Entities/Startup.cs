namespace WordProcessor.Table1.Entities;

public class Startup
{
    public string Number { get; set; }   
    public string Name { get; set; }
    public string Link { get; set; }
    
    public string DupeCount { get; set; }
    public string HasSign { get; set; }
    public string Category { get; set; }
    public List<Participant> Participants { get; set; }
}