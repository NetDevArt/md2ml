using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;
using Md2Ml.Enum;

namespace Md2Ml
{
    internal class MarkdownStringParser
    {
        /// <summary>
		/// Parse the content, and detect all real lines breaks.
		/// In fact, within markdown it is possible to write a unique paragraph, with line breaks.
		/// If there is not a double space before the line break, it's the same paragraph, and the line break is not interpreted as is.
		///
		/// So in this parsing method, cut the content by "block" with same nature.
		///		1| Detect the paragraph type by the start of the line
		///		2| Process the content by its type
		///		3| Then remove the processed content, and continue to parse the next content
		///
        /// A block is detected by a <see cref="ParaPattern"/> which have an associated regex.
		/// 
		/// </summary>
		/// <param name="engine">Describe an openXML object by code</param>
		/// <param name="mdText">The markdown text to parse</param>
		internal static void Parse(Md2MlEngine engine, string mdText)
        {
            while (!string.IsNullOrEmpty(mdText))
            {
                var firstLine = GetFirstLine(mdText);
                var matchedPattern = PatternMatcher.GetMarkdownMatch(firstLine);
                Paragraph para;
                (int counter, string textBlock) rebuildText = default;
                switch (matchedPattern.Key)
                {
					case ParaPattern.InfiniteHeading:
                        string titleChars = matchedPattern.Value.Groups[1].Value;
						int titleLvl = titleChars.Count(c => c == '#');
                        titleLvl = titleLvl <= 9 ? titleLvl : 9;
						para = engine.CreateParagraph(new ParaProperties() { StyleName = string.Concat("Heading", titleLvl.ToString()) });
                        engine.WriteText(para, matchedPattern.Value.Groups[2].Value);
                        mdText = DeleteLines(mdText);
                        continue;
                    case ParaPattern.Image:
                        var link = matchedPattern.Value.Groups[2].Value;
                        if (link.StartsWith("http://") || link.StartsWith("https://"))
                            engine.AddImage(new System.Net.WebClient().OpenRead(link));
                        else
                            engine.AddImage(ConvertRelativeToAbsolutePath(link, engine.GetFileDirectory()));
                        mdText = DeleteLines(mdText);
                        continue;
                    case ParaPattern.Table:
                    //case ParaPattern.TableHeaderSeparation:
                        rebuildText = BuildWithoutBreakingLines(mdText, ParaPattern.Table);
                        ProcessTable(engine, rebuildText.textBlock);
                        mdText = DeleteLines(mdText, rebuildText.counter);
                        continue;
                    case ParaPattern.OrderedList:
                    case ParaPattern.UnorderedList:
                        // TODO : Ordered list can contain unordered items inside and vice versa
                        rebuildText = BuildWithoutBreakingLines(mdText, ParaPattern.OrderedList);
                        ProcessBullets(engine, rebuildText.textBlock);
                        mdText = DeleteLines(mdText, rebuildText.counter);
                        continue;
                    case ParaPattern.CodeBlock:
                        // TODO : Improve the rendering - not a priority for my needs
                        rebuildText = BuildWithoutBreakingLines(mdText, ParaPattern.CodeBlock);
                        para = engine.CreateParagraph(new ParaProperties() { StyleName = DocStyles.CodeBlock.ToDescriptionString() });
                        FormatText(engine, para, rebuildText.textBlock, new StyleProperties());
                        mdText = DeleteLines(mdText, rebuildText.counter);
                        continue;
                    case ParaPattern.Quote:
                        // Markdown supports nested quotes, but word does not
                        // So whatever, put the "nested" paragraph in the same quote
                        rebuildText = BuildWithoutBreakingLines(mdText, ParaPattern.Quote);
                        para = engine.CreateParagraph(new ParaProperties() { StyleName = DocStyles.Quote.ToDescriptionString()});
                        foreach (var text in Regex.Split(rebuildText.textBlock, "\r\n|\r|\n"))
                        {
                            para.AppendChild(new Break());
                            FormatText(engine, para, text, new StyleProperties());
                        }
                        mdText = DeleteLines(mdText, rebuildText.counter);
                        continue;
                    case ParaPattern.AnyChar:
                    default:
                        rebuildText = BuildWithoutBreakingLines(mdText, matchedPattern.Key);
                        para = engine.CreateParagraph();
                        FormatText(engine, para, rebuildText.textBlock, new StyleProperties());
                        // engine.WriteText(para, text.textBlock);
                        mdText = DeleteLines(mdText, rebuildText.counter);
                        continue;
				}
            }
        }

        /// <summary>
        /// Process a full string block of a markdown table.
        /// Split it by rows
        /// Create the Table with correct number of columns
        /// Fill the openXML table
        /// </summary>
        /// <param name="engine"></param>
        /// <param name="markdown">The complete markdown table, must be correctly formatted</param>
        private static void ProcessTable(Md2MlEngine engine, string markdown)
        {
            var rows = Regex.Split(markdown, "\r\n|\r|\n");
            var firstLine = rows.First();
            var secondLine = rows.Length >= 2 ? rows.Skip(1).First() : null;
            var table = engine.CreateTable(firstLine.Trim('|').Split('|').Count());
            engine.AddTableRow(table, firstLine.Trim('|').Split('|').ToList());
            var patternSecondLine = PatternMatcher.GetMarkdownMatch(secondLine).Key;
            if (string.IsNullOrEmpty(secondLine) || 
                (patternSecondLine != ParaPattern.TableHeaderSeparation && patternSecondLine != ParaPattern.TableHeaderSeparation)) // TODO : Throw error : Table not well formatted
                return;

            // Define the table alignment properties
            List<JustificationValues> cellJustification = new List<JustificationValues>();
            var nbCols = secondLine.Trim('|').Split('|').Count();
            var secondLineCells = secondLine.Trim('|').Split('|').ToList();
            for (int i = 0; i < nbCols; i++)
            {
                var justification = JustificationValues.Left;
                if(secondLineCells[i].StartsWith(":") && secondLineCells[i].EndsWith(":"))
                    justification = JustificationValues.Center;
                else if (!secondLineCells[i].StartsWith(":") && secondLineCells[i].EndsWith(":"))
                    justification = JustificationValues.Right;

                cellJustification.Add(justification);
            }

            // Process the rest of the table 
            foreach (var row in rows.Skip(2).ToList())
                engine.AddTableRow(table, row.Trim('|').Split('|').ToList(), cellJustification);
        }

        private static void ProcessBullets(Md2MlEngine engine, string markdown)
        {
            // Split the text block into a list of items
            var lines = Regex.Split(markdown, "\r\n|\r|\n").ToArray();
            engine.MarkdownList(engine, lines);
        }


        /// <summary>
        /// Format a paragraph which could contains some text styles, Bold, Italics, Images and so on...
        /// Split the markdown for each pattern found. Then append correctly that text or image into the same paragraph.
        /// </summary>
        /// <param name="core">The openXML object with a document, a body and a paragraph</param>
        /// <param name="paragraph">The Paragraph object previously created</param>
        /// <param name="markdown">The string to be processed</param>
        /// <param name="fontProperties">Style properties to apply to the text</param>
        internal static void FormatText(Md2MlEngine core, Paragraph paragraph, string markdown,
            StyleProperties fontProperties)
        {
            var hasPattern = PatternMatcher.HasPatterns(markdown);
            while (hasPattern)
            {
                var s = PatternMatcher.GetPatternsAndNonPatternText(markdown);
                var newFontProperties = new StyleProperties();
                
                switch (s.Key)
                {
                    case StylePattern.BoldAndItalic:
                        newFontProperties.Bold = true;
                        newFontProperties.Italic = true;
                        FormatText(core, paragraph, s.Value[0], new StyleProperties());
                        FormatText(core, paragraph, s.Value[1], newFontProperties);
                        FormatText(core, paragraph, FramePendingString(s.Value, "***"), new StyleProperties());
                        break;
                    case StylePattern.Bold:
                        newFontProperties.Bold = true;
                        FormatText(core, paragraph, s.Value[0], new StyleProperties());
                        FormatText(core, paragraph, s.Value[1], newFontProperties);
                        FormatText(core, paragraph, FramePendingString(s.Value, "**"), new StyleProperties());
                        break;
                    case StylePattern.Italic:
                        newFontProperties.Italic = true;
                        FormatText(core, paragraph, s.Value[0], new StyleProperties());
                        FormatText(core, paragraph, s.Value[1], newFontProperties);
                        FormatText(core, paragraph, FramePendingString(s.Value, "*"), new StyleProperties());
                        break;
                    case StylePattern.MonospaceOrCode:
                        newFontProperties.StyleName = DocStyles.CodeReference.ToDescriptionString();
                        FormatText(core, paragraph, s.Value[0], new StyleProperties());
                        FormatText(core, paragraph, s.Value[1], newFontProperties);
                        FormatText(core, paragraph, FramePendingString(s.Value, "`"), new StyleProperties());
                        break;
                    case StylePattern.Strikethrough:
                        newFontProperties.Strikeout = true;
                        FormatText(core, paragraph, s.Value[0], new StyleProperties());
                        FormatText(core, paragraph, s.Value[1], newFontProperties);
                        FormatText(core, paragraph, FramePendingString(s.Value, "~~"), new StyleProperties());
                        break;
                    case StylePattern.Image:
                        var regex = PatternMatcher.GetStyleMatch(s.Value[1]);
                        FormatText(core, paragraph, s.Value[0], new StyleProperties());
                        core.AddImage(ConvertRelativeToAbsolutePath(regex.Value.Groups[2].Value, core.GetFileDirectory()), paragraph);
                        FormatText(core, paragraph, FramePendingString(s.Value, ""), new StyleProperties());
                        break;
                    case StylePattern.Underline:
                        newFontProperties.Underline = UnderlineValues.Single;
                        FormatText(core, paragraph, s.Value[0], new StyleProperties());
                        FormatText(core, paragraph, s.Value[1], newFontProperties);
                        FormatText(core, paragraph, FramePendingString(s.Value, "__"), new StyleProperties());
                        break;
                }
                return;
            }
            core.WriteText(paragraph, markdown, fontProperties);
        }
        
        /// <summary>
        /// Concatenate the rest of the string, replacing the pattern around matched string
        /// And remove the previously processed string
        /// </summary>
        /// <param name="strs"></param>
        /// <param name="patten"></param>
        /// <param name="startConcatIndex"></param>
        /// <returns></returns>
        private static string FramePendingString(IReadOnlyList<string> strs, string patten, int startConcatIndex = 2)
        {
            var str = "";
            for (int i = startConcatIndex; i < strs.Count(); i++)
            {
                if (i % 2 == 0)
                    str += strs[i];
                else
                    str += patten + strs[i] + patten;
            }
            return str;
        }

        /// <summary>
        /// Remove lines after processing it
        /// </summary>
        /// <param name="mdText">The text where delete lines from top</param>
        /// <param name="nbLine">The number of lines to delete</param>
        /// <returns></returns>
        private static string DeleteLines(string mdText, int nbLine = 1)
        {
            var lines = Regex.Split(mdText, "\r\n|\r|\n").Skip(nbLine).ToArray();
            return string.Join(Environment.NewLine, lines);
        }
        private static string GetFirstLine(string mdText)
        {
            return Regex.Split(mdText, "\r\n|\r|\n").First();
        }

        /// <summary>
        /// This method allows you to rebuild a paragraph without interpreting breaking lines if there is not a double space before.
        /// Produces a counter in order to know how many lines are concatenated, in order to delete it after parsing it
        /// </summary>
        /// <param name="mdText">The text to parse line by line</param>
        /// <param name="pattern">The pattern to detect real new lines. Allows to create a block with same nature (lists, paragraphs...)</param>
        /// <returns></returns>
        private static (int counter, string textBlock) BuildWithoutBreakingLines(string mdText, ParaPattern pattern)
        {
            var lines = Regex.Split(mdText, "\r\n|\r|\n").ToArray();
            var previousLine = lines.First().TrimStart('>');
            var output = new StringBuilder(previousLine);
            int count = 1;
            if (string.IsNullOrEmpty(previousLine)) return (counter: count, textBlock: output.ToString());

            foreach (var line in lines.Skip(1))
            {
                // Break directly if processed line is empty
                if (string.IsNullOrEmpty(line)) break;
                var linePattern = PatternMatcher.GetMarkdownMatch(line).Key;

                // If first pattern is a table, do not break until pattern does not match to any table pattern
                bool isTableContinuing = pattern == ParaPattern.Table && (
                    linePattern == ParaPattern.Table ||
                    linePattern == ParaPattern.TableHeaderSeparation);

                // Quotes are continuing if first pattern is quote
                // And the current one equals to Quote or AnyChar
                bool isQuoteContinuing = pattern == ParaPattern.Quote &&
                                         (linePattern == ParaPattern.Quote ||
                                          linePattern == ParaPattern.AnyChar);

                // Paragraph types are continuing if previous line does not ends with double space
                // And first pattern matches with the current one
                bool isParagraphContinuing = (!previousLine.EndsWith("  ") &&
                                            (pattern == ParaPattern.AnyChar && linePattern == pattern));

                // If list is not continuing, then check if item is detected as a code block
                bool isListContinuing = (pattern == ParaPattern.OrderedList || pattern == ParaPattern.UnorderedList) &&
                                        (linePattern == ParaPattern.OrderedList ||
                                         linePattern == ParaPattern.UnorderedList);
                if ((pattern == ParaPattern.OrderedList || pattern == ParaPattern.UnorderedList) && !isListContinuing && linePattern == ParaPattern.CodeBlock)
                {
                    try
                    {
                        linePattern = PatternMatcher.GetMatchFromPattern(line, ParaPattern.OrderedList).Key;
                    }
                    catch (Exception e)
                    {
                        linePattern = PatternMatcher.GetMatchFromPattern(line, ParaPattern.UnorderedList).Key;
                    }

                    isListContinuing = (linePattern == ParaPattern.OrderedList || linePattern == ParaPattern.UnorderedList);
                }


                if (!isTableContinuing && !isQuoteContinuing && !isParagraphContinuing && !isListContinuing) break;

                
                if (pattern == ParaPattern.TableHeaderSeparation ||
                    pattern == ParaPattern.Table ||
                    pattern == ParaPattern.OrderedList ||
                    pattern == ParaPattern.UnorderedList ||
                    pattern == ParaPattern.Quote)
                    output.AppendLine().Append(line.TrimStart('>'));
                else
                    output.Append(line);

                previousLine = line;
                count++;
            }
            return (counter: count, textBlock: output.ToString());
        }

        /// <summary>
        /// This method allows you to convert a relative image path to an absolute.
        /// Be sure to set the directory path of the parsed file. <see cref="Md2MlEngine.SetFileDirectory"/>
        /// </summary>
        /// <param name="imgPath"></param>
        /// <param name="fileDir"></param>
        /// <returns></returns>
        private static string ConvertRelativeToAbsolutePath(string imgPath, string fileDir = null)
        {
            if (File.Exists(imgPath)) return imgPath;
            var absolutePath = imgPath;
            if (!string.IsNullOrEmpty(fileDir))
                absolutePath = Path.Combine(fileDir, imgPath);
            // TODO : Reconsider removing the throwing exception if file does not exists
            if (!File.Exists(absolutePath)) throw new IOException($"The file {absolutePath} does not exists");
            return absolutePath;
        }

    }
}
