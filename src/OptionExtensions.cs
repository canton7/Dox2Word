using System;
using System.Linq;
using Mono.Options;

namespace Dox2Word
{
    public static class OptionExtensions
    {
        public static OptionSet AddRequired(this OptionSet optionSet, string prototype, string description, Action<string> action) =>
            optionSet.AddRequired<string>(prototype, description, action);

        public static OptionSet AddRequired<T>(this OptionSet optionSet, string prototype, string description, Action<T> action)
        {
            return optionSet.Add(new CustomOptionOption<T>(prototype, description, action, isRequired: true));
        }

        public static void ValidateRequiredOptions(this OptionSet optionSet)
        {
            foreach (var option in optionSet.OfType<IRequiredOption>())
            {
                option.ValidateRequired();
            }
        }

        private interface IRequiredOption
        {
            void ValidateRequired();
        }

        private class CustomOptionOption<T> : Option, IRequiredOption
        {
            private readonly Action<T> action;
            private readonly bool isRequired;
            private bool isParsed;

            public CustomOptionOption(string prototype, string description, Action<T> action, bool isRequired)
                    : base(prototype, description, 1)
            {
                if (action == null)
                    throw new ArgumentNullException("action");
                this.action = action;
                this.isRequired = isRequired;
            }

            protected override void OnParseComplete(OptionContext c)
            {
                if (string.IsNullOrEmpty(c.OptionValues[0]))
                    throw new OptionException(string.Format(c.OptionSet.MessageLocalizer("Missing required value for option '{0}'."), c.OptionName), c.OptionName);

                this.action(Parse<T>(c.OptionValues[0], c));
                this.isParsed = true;
            }

            public void ValidateRequired()
            {
                if (this.isRequired && !this.isParsed)
                {
                    throw new OptionException($"Missing required option '-{this.GetNames()[0]}'", this.GetNames()[0]);
                }
            }
        }
    }
}
