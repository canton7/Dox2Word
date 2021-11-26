namespace Dox2Word.Model
{
    internal interface IMergable<T>
    {
        string Id { get; }
        void MergeWith(T other);
    }
}
