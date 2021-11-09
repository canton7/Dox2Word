namespace Dox2Word.Model
{
    public class VariableDoc : IMergable<VariableDoc>
    {
        public string Id { get; set; } = null!;
        public string? Name { get; set; }
        public string? Type { get; set; }
        public string? Definition { get; set; }
        public Descriptions Descriptions { get; set; } = null!;
        public string? Initializer { get; internal set; }
        public string? Bitfield { get; set; }

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
            if (other.Initializer == null)
            {
                this.Initializer = null;
            }
            this.Bitfield ??= other.Bitfield;

        }
    }
}
