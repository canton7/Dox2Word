using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dox2Word.Model
{
    public enum ParagraphType
    {
        Normal,
        Warning,
    }

    public class Paragraph : Collection<ITextRun>
    {
        public ParagraphType Type { get; }

        public Paragraph(ParagraphType type = ParagraphType.Normal)
        {
            this.Type = type;
        }
    }
}
