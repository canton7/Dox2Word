using System.Collections.Generic;
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
        private readonly HashSet<string> namesInUse = new();
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

        public Hyperlink CreateLink(string name, Run text)
        {
            if (!this.names.ContainsKey(name))
            {
                this.referencedButNotCreated.Add(name);
            }

            var hyperlink = new Hyperlink(text)
            {
                Anchor = this.TransformName(name),
                History = true,
            };
            return hyperlink;
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
            string transformedName;
            // Have we transformed this one before? Use the result of that?
            if (!this.names.TryGetValue(name, out transformedName))
            {
                if (name.Length <= MaxLength)
                {
                    this.names[name] = name;
                    this.namesInUse.Add(name);
                    transformedName = name;
                }
                else
                {
                    for (int i = 1; ; i++)
                    {
                        string iString = i.ToString();
                        string renamed = name.Substring(0, MaxLength - (1 + iString.Length)) + "+" + iString;
                        if (!this.namesInUse.Contains(renamed))
                        {
                            this.names[name] = renamed;
                            this.namesInUse.Add(renamed);
                            transformedName = renamed;
                            break;
                        }
                    }
                }
            }

            return transformedName;
        }
    }
}
