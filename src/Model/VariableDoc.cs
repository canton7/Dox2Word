using System.Collections.Generic;

namespace Dox2Word.Model
{
    public class VariableDoc : IMergable<VariableDoc>
    {
        public string Id { get; set; } = null!;
        public string? Name { get; set; }
        public List<TextRun> Type { get; set; } = null!;
        public string? Definition { get; set; }
        public Descriptions Descriptions { get; set; } = null!;
        public List<TextRun> Initializer { get; set; } = null!;
        public string? Bitfield { get; set; }
        public string? ArgsString { get; set; }

        void IMergable<VariableDoc>.MergeWith(VariableDoc other)
        {
            // We can have multiple definitions of a variable, e.g. with an extern if one part
            // is in the header and another in the body.
            // They should have identical information (but pool it anyway), but if either has no
            // initializer neither should (as the initializer is probably in the C file, and we don't
            // care about it).

            this.Name ??= other.Name;
            this.Type ??= other.Type;
            this.Definition ??= other.Definition;
            this.Descriptions.MergeWith(other.Descriptions);
            if (other.Initializer.Count == 0)
            {
                this.Initializer.Clear();
            }
            this.Bitfield ??= other.Bitfield;
            this.ArgsString ??= other.ArgsString;
        }
    }
}
