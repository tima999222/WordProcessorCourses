namespace WordProcessor.Data.Repositrory
{
    public interface IBaseRepository<T>
    {
        IEnumerable<T> GetListOfItems();
    }
}
