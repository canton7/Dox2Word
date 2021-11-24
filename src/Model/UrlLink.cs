using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Dox2Word.Model
{
    public class UrlLink : Collection<TextRun>, IParagraphElement
    {
        public string Url { get; }
        public UrlLink(string url)
        {
            this.Url = url;
        }

        public void AddRange(IEnumerable<TextRun> elements)
        {
            foreach (var element in elements)
            {
                this.Add(element);
            }
        }
    }
}
