using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dox2Word.Model
{
    internal interface IMergable<T>
    {
        string Id { get; }
        void MergeWith(T other);
    }
}
