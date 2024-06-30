namespace WordProcessor
{
    public interface IWordGenerator
    {
        void GenerateWord<T>(string pathTemplate, string pathDestination, T obj);
    }
}
