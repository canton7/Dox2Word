using System.Collections.Generic;

namespace Dox2Word.Model
{
    public class ProjectOption
    {
        public string Id { get; set; } = null!;
        public List<string> Values { get; } = [];
        public object TypedValue { get; set; } = null!;
    }
}