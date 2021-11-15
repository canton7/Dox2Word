using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dox2Word.Generator
{
    public static class FieldReplacer
    {
        public static void Replace(OpenXmlElement element, Dictionary<string, string> substitutions)
        {
            foreach (var match in FindPlaceholders(element))
            {
                if (substitutions.TryGetValue(match.Text, out string replacement))
                {

                }
            }
        }

        public static void Replace(OpenXmlElement element, string placeholder, Paragraph replacement)
        {
            var match = FindPlaceholders(element).FirstOrDefault(x => x.Text == placeholder);
            if (match.Text != null)
            {
                if (match.Start == match.End)
                {

                }
                else
                {
                    match.Start.Text = match.Start.Text.Substring(match.StartPos);
                    if (string.IsNullOrEmpty(match.Start.Text))
                    {
                        match.Start.Remove();
                    }
                    match.End.Text = match.End.Text.Substring(0, match.EndPos);
                    if (string.IsNullOrEmpty(match.End.Text))
                    {
                        match.End.Remove();
                    }
                }
            }
        }

        private static void RemovePlaceholder(Placeholder placeholder)
        {
            foreach (var member in placeholder.Members)
            {
                member.Remove();
                if (member.Parent!.ChildElements.Count == 0)
                {
                    member.Parent.Remove();
                }
            }

            if (placeholder.Start != placeholder.End)
            {
                placeholder.End.Text = placeholder.End.Text.Substring(placeholder.EndPos);
            }
        }

        private static IEnumerable<Placeholder> FindPlaceholders(OpenXmlElement element)
        {
            var sb = new StringBuilder();
            var members = new List<Text>();

            foreach (var paragraph in element.Descendants<Paragraph>())
            {
                Text? start = null;

                foreach (var text in paragraph.Descendants<Text>())
                {
                    int index = -1;
                    int startIndex = -1;

                    while (true)
                    {
                        index = text.Text.IndexOfAny(new[] { '<', '>' }, index + 1);
                        if (index == -1)
                        {
                            if (start != null)
                            {
                                if (start != text)
                                {
                                    members.Add(text);
                                }
                                sb.Append(text.Text.Substring(startIndex + 1));
                            }
                            break;
                        }
                        else if (text.Text[index] == '<')
                        {
                            start = text;
                            startIndex = index;
                            sb.Clear();
                            members.Clear();
                        }
                        else if (start != null)
                        {
                            sb.Append(text.Text.Substring(startIndex + 1, index - startIndex - 1));
                            yield return new Placeholder(sb.ToString(), start, startIndex, members, text, index);
                            start = null;
                        }
                    }
                }
            }
        }

        private struct Placeholder
        { 
            public string Text { get; }
            public Text Start { get; }
            public int StartPos { get; }
            public List<Text> Members { get; }
            public Text End { get; }
            public int EndPos { get; }

            public Placeholder(string placeholder, Text start, int startPos, List<Text> members, Text end, int endPos)
            {
                this.Text = placeholder;
                this.Start = start;
                this.StartPos = startPos;
                this.Members = members;
                this.End = end;
                this.EndPos = endPos;
            }
        }
    }
}
