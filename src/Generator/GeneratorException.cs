using System;

namespace Dox2Word.Generator
{
    internal class GeneratorException : Exception
    {
        public GeneratorException()
        {
        }

        public GeneratorException(string message) : base(message)
        {
        }

        public GeneratorException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}