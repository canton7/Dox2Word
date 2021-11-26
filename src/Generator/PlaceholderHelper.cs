using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Dox2Word.Logging;

namespace Dox2Word.Generator
{
    public static class PlaceholderHelper
    {
        private static readonly Logger logger = Logger.Instance;

        private delegate (string? replacement, bool abort) PlaceholderFoundAction(Text text, string placeholder);

        public static void Replace(OpenXmlElement element, Dictionary<string, string> substitutions)
        {
            FindPlaceholders(element, (text, placeholder) =>
            {
                if (substitutions.TryGetValue(placeholder, out string? replacement))
                {
                    logger.Info($"Replacing placeholder '{placeholder}' with '{replacement}'");
                }
                return (replacement, false);
            });
        }

        public static Paragraph? FindParagraphPlaceholder(OpenXmlElement element, string placeholder)
        {
            Paragraph? result = null;
            FindPlaceholders(element, (text, foundPlaceholder) =>
            {
                if (placeholder == foundPlaceholder)
                {
                    result = (Paragraph)text.Parent!.Parent!;
                    return (null, true);
                }
                return (null, false);
            });
            return result;
        }

        private static void FindPlaceholders(OpenXmlElement element, PlaceholderFoundAction foundAction)
        {
            var sb = new StringBuilder();
            var members = new List<Text>();

            foreach (var paragraph in element.Descendants<Paragraph>())
            {
                Text? start = null;
                int indexOfStartInStart = -1;

                // TODO: We probably want to ignore Space elements
                foreach (var text in paragraph.Descendants<Text>())
                {
                    int index = -1;
                    int indexOfStartInCurrentText = -1;

                    while (true)
                    {
                        if (string.IsNullOrEmpty(text.Text))
                            break;

                        index = text.Text.IndexOfAny(new[] { '<', '>' }, index + 1);
                        if (index == -1)
                        {
                            if (start != null)
                            {
                                if (start != text)
                                {
                                    members.Add(text);
                                }
                                sb.Append(text.Text.Substring(indexOfStartInCurrentText + 1));
                            }
                            break;
                        }
                        else if (text.Text[index] == '<')
                        {
                            start = text;
                            indexOfStartInCurrentText = index;
                            indexOfStartInStart = index;
                            sb.Clear();
                            members.Clear();
                        }
                        else if (start != null) // text.Text[index] == '>'
                        {
                            sb.Append(text.Text.Substring(indexOfStartInCurrentText + 1, index - indexOfStartInCurrentText - 1));
                            var (replacement, abort) = foundAction(start, sb.ToString());
                            if (abort)
                            {
                                return;
                            }

                            if (replacement != null)
                            {
                                foreach (var member in members)
                                {
                                    var parent = member.Parent;
                                    member.Remove();
                                    if (parent!.ChildElements.Count == 0)
                                    {
                                        parent.Remove();
                                    }
                                }

                                if (start == text)
                                {
                                    string original = start.Text;
                                    start.Text = original.Substring(0, indexOfStartInStart)
                                        + replacement + original.Substring(index + 1);
                                    index = indexOfStartInStart + replacement.Length - 1;
                                }
                                else
                                {
                                    start.Text = start.Text.Substring(0, indexOfStartInStart) + replacement;
                                    text.Text = text.Text.Substring(index + 1);
                                    index = 0;
                                }

                                start = null;
                            }
                        }
                    }
                }
            }
        }
    }
}
