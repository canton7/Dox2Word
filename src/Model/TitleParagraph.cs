using System.Collections.ObjectModel;
using System.Linq;

namespace Dox2Word.Model
{
    public class TitleParagraph : Collection<IParagraphElement>, IParagraph
    {
        public bool IsEmpty => this.Count == 0;

        public void TrimTrailingWhitespace()
        {
            if (this.LastOrDefault() is TextRun run && run.Text.EndsWith(" "))
            {
                run.Text = run.Text.TrimEnd(' ');
                if (run.Text.Length == 0)
                {
                    this.Remove(run);
                }
            }
        }
    }
}
