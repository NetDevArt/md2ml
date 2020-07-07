using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text.RegularExpressions;
using Md2Ml.Enum;

namespace Md2Ml
{
    class PatternMatcher
    {
        /// <summary>
        /// Regular expressions to detect patterns of markdown elements at the LINE START
        /// It means if a paragraph contains multiple lines, only the first one defines the type
        /// Beware that some regex have multiple groups 
        /// </summary>
        private static Dictionary<ParaPattern, Regex> ParagraphPatterns = new Dictionary<ParaPattern, Regex>()
        {
            /**
             * Detect headings, even those higher than 6, 2 groups:
             *      1| All '#' enven if there is spaces before text
             *      2| The content of the title
             */
            { ParaPattern.InfiniteHeading, new Regex(@"^(#+[\s|#]+)(.*)", RegexOptions.Multiline | RegexOptions.Compiled) },
            
            /**
             * Detect code block, 2 groups :
             *  1| A tab or a space char x4
             *  2| The content
            */
            { ParaPattern.CodeBlock, new Regex(@"^([ ]{4}|^\t{1})(.*)", RegexOptions.Multiline | RegexOptions.Compiled) },

            /**
             * Detect images (only one image on the line), 2 groups:
             *      1| The title of the image
             *      2| The path to the image
             */
            { ParaPattern.Image, new Regex(@"^!\[([\w|\s]*)\]\(([\w|\S]*)\)[ ]*$", RegexOptions.Multiline | RegexOptions.Compiled)},
            
            /**
             *  Detect Ordered list, 2 groups:
             *      1| The number before the point
             *      2| The content after the point
             */
            { ParaPattern.OrderedList, new Regex(@"^([ ]{3})*\d+\. (.*)", RegexOptions.Multiline | RegexOptions.Compiled)},
            
            /**
             * Detect Unordered List, 2 groups:
             *      1| How many spaces there are (could be null if top level list)
             *      2| The content after the list charachter
             */
            { ParaPattern.UnorderedList, new Regex(@"^([ ]{3})*[*+-] (.*)", RegexOptions.Multiline | RegexOptions.Compiled) },

            /**
             * Detect a (nested too) Quote  block, without forcing the space after the ">", 2 groups:
             *      1| How many '>' there are
             *      2| The content of the citation
             */
            { ParaPattern.Quote, new Regex(@"^(>+)(.*)", RegexOptions.Compiled) },
            { ParaPattern.TableHeaderSeparation, new Regex(@"^(\|\W+\|)", RegexOptions.Compiled)},
            { ParaPattern.Table, new Regex(@"\|(.*)\|", RegexOptions.Compiled) },
            

            /**
             *  Detect requirement formatted (special format described by THALES)
             */
            {ParaPattern.ReqTitle, new Regex(@"^(\[\w+-\w+-REQ-\d+])([\w|\s]+)", RegexOptions.Compiled)},
            {ParaPattern.ReqProperties1, new Regex(@"^(@[\w|\s]+:)(.+)", RegexOptions.Compiled)},
            {ParaPattern.ReqProperties2, new Regex(@"^(%[\w|\s]+:)(.+)", RegexOptions.Compiled)},
            
            /**
             * Any char except line terminator
             */
            { ParaPattern.AnyChar, new Regex("(.*)", RegexOptions.Multiline | RegexOptions.Compiled)}
        };

        /// <summary>
        /// Regular expression to detect markdown styling elements within a paragraph.
        /// Basically, which elements are contained in a paragraph
        /// </summary>
        private static Dictionary<StylePattern, Regex> StylePatterns = new Dictionary<StylePattern, Regex>()
        {
            { StylePattern.Bold, new Regex(@"(?<!\*)\*\*([^\*].+?)\*\*") },
            { StylePattern.Italic, new Regex(@"(?<!\*)\*([^\*].+?)\*") },
            { StylePattern.BoldAndItalic, new Regex(@"(?<!\*)\*\*\*([^\*].+?)\*\*\*") },
            { StylePattern.Image, new Regex(@"!\[([\w|\s]*)\]\(([\w|\S]*)\)", RegexOptions.Multiline | RegexOptions.Compiled)},
            { StylePattern.Link, new Regex(@"\[(.+?)\]\((.+)\)") },
            { StylePattern.MonospaceOrCode, new Regex(@"`{1}([^`]+)`{1}") },
            { StylePattern.Strikethrough, new Regex(@"~{2}(.*)~{2,}") },
            { StylePattern.Tab, new Regex(@"\t(.*)") },
            { StylePattern.Underline, new Regex(@"_{2}(.*)_{2,}") }
        };

        /// <summary>
        /// Check the type of the markdown string in params, and return the match
        /// </summary>
        /// <param name="markdown">The string to analyze</param>
        /// <returns>A KeyValuePair with pattern id and the match</returns>
        public static KeyValuePair<ParaPattern, Match> GetMarkdownMatch(string markdown)
        {
            foreach (var pattern in ParagraphPatterns)
            {
                var regex = pattern.Value;
                if (!regex.IsMatch(markdown)) continue;
                return new KeyValuePair<ParaPattern, Match>(pattern.Key, regex.Match(markdown));
            }
            throw new NotSupportedException();
        }

        public static KeyValuePair<ParaPattern, Match> GetMatchFromPattern(string markdown, ParaPattern pattern)
        {
            var regex = ParagraphPatterns[pattern];
            if (!regex.IsMatch(markdown)) throw new AmbiguousMatchException("The match does not refer to the needed.");
            return new KeyValuePair<ParaPattern, Match>(pattern, regex.Match(markdown));
        }

        public static KeyValuePair<StylePattern, Match> GetStyleMatch(string markdown)
        {
            foreach (var pattern in StylePatterns)
            {
                var regex = pattern.Value;
                if (!regex.IsMatch(markdown)) continue;
                return new KeyValuePair<StylePattern, Match>(pattern.Key, regex.Match(markdown));
            }
            throw new NotSupportedException();
        }

        public static bool HasPatterns(string markdown)
        {
            foreach (var pattern in StylePatterns)
            {
                var regex = pattern.Value;
                if (!regex.IsMatch(markdown)) continue;
                return true;
            }
            return false;
        }

        public static KeyValuePair<StylePattern, string[]> GetPatternsAndNonPatternText(string markdown)
        {

            foreach (var pattern in StylePatterns)
            {
                var regex = pattern.Value;
                if (!regex.IsMatch(markdown)) continue;
                // Indexes
                var startSplitting = regex.Match(markdown).Index;
                var splitLength = regex.Match(markdown).Length;
                // Substrings
                var before = markdown.Substring(0, startSplitting);
                var match = markdown.Substring(startSplitting, splitLength);
                if (pattern.Key != StylePattern.Image)
                    match = regex.Split(match)[1];
                var after = markdown.Substring(startSplitting + splitLength, markdown.Length - (startSplitting + splitLength));
                return new KeyValuePair<StylePattern, string[]>(pattern.Key, new []{before, match, after});
            }
            return new KeyValuePair<StylePattern, string[]>(StylePattern.PlainText, new string[] { markdown });
        }
    }
}
