using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Dox2Word.Logging;

namespace Dox2Word.Generator
{
    public class BookmarkManager
    {
        private static readonly Logger logger = Logger.Instance;

        private const int MaxLength = 35;

        // Original, requested name -> name we had to use
        private readonly Dictionary<string, string> names = new();
        private readonly HashSet<string> referencedButNotCreated = new();

        private int id;

        public BookmarkManager(Body preExistingDocument)
        {
            int maxId = preExistingDocument.Descendants<BookmarkStart>().MaxOrDefault(x => int.Parse(x.Id), -1);
            this.id = maxId + 1;
        }

        public (BookmarkStart start, BookmarkEnd end) CreateBookmark(string name)
        {
            this.referencedButNotCreated.Remove(name);
            var result = (new BookmarkStart() { Id = this.id.ToString(), Name = this.TransformName(name) }, new BookmarkEnd() { Id = this.id.ToString() });
            this.id++;
            return result;
        }

        public Run[] CreateLink(string name, Run text)
        {
            if (!this.names.ContainsKey(name))
            {
                this.referencedButNotCreated.Add(name);
            }

            var runs = new[]
            {
                new Run(new FieldChar()
                {
                    FieldCharType = FieldCharValues.Begin,
                }),
                new Run(new FieldCode()
                {
                    Text = $" REF {this.TransformName(name)} \\h ",
                    Space = SpaceProcessingModeValues.Preserve,
                }),
                new Run(new FieldCode()
                {
                    Text = "\\* MERGEFORMAT ",
                    Space = SpaceProcessingModeValues.Preserve,
                }),
                new Run(new FieldChar()
                {
                    FieldCharType = FieldCharValues.Separate,
                }),
                text,
                new Run(new FieldChar()
                {
                    FieldCharType = FieldCharValues.End,
                })
            };
            return runs;
        }

        public void Validate()
        {
            foreach (string notCreated in this.referencedButNotCreated)
            {
                logger.Warning($"Bookmark {notCreated} was referenced but not created");
            }
        }

        private string TransformName(string name)
        {
            string actualName;
            if (!this.names.TryGetValue(name, out actualName))
            {
                if (name.Length <= MaxLength)
                {
                    this.names[name] = name;
                    actualName = name;
                }
                else
                {
                    // Try just truncating it
                    string truncated = name.Substring(0, MaxLength);
                    if (!this.names.ContainsKey(truncated))
                    {
                        this.names[name] = truncated;
                        actualName = truncated;
                    }
                    else
                    {
                        // That didn't work. 
                        for (int i = 1; ; i++)
                        {
                            string renamed = name.Substring(0, MaxLength - 10) + i;
                            if (!this.names.ContainsKey(renamed))
                            {
                                this.names[name] = renamed;
                                actualName = renamed;
                                break;
                            }
                        }
                    }
                }
            }

            return actualName;
        }
    }
}
