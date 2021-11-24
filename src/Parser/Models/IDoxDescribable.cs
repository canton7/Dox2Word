namespace Dox2Word.Parser.Models
{
    public interface IDoxDescribable
    {
        Description? BriefDescription { get; }
        Description? DetailedDescription { get; }
    }
}
