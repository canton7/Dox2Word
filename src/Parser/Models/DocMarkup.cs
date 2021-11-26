using Dox2Word.Model;

namespace Dox2Word.Parser.Models
{
    public class DocMarkup : DocCmdGroup
    {
        public TextRunProperties Properties { get; set; }

        public bool Center { get; set; }
    }
}
