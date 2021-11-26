namespace Dox2Word.Model
{
    public interface IParagraph
    {
        bool IsEmpty { get; }

        void TrimTrailingWhitespace();
    }
}
