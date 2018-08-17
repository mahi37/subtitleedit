using System;
using System.Collections.Generic;
using System.Text;

namespace Nikse.SubtitleEdit.Core.SubtitleFormats
{
    /// <summary>
    /// CBTNuggets, e.g. https://portal.cbtnuggets.com/mediav2/caption/54f4cf34fcb8f7d56e0001bf_10x.json
    /// </summary>
    public class JsonType9 : SubtitleFormat
    {
        public override string Extension => ".json";

        public override string Name => "JSON Type 9";

        public override string ToText(Subtitle subtitle, string title)
        {
            var sb = new StringBuilder(@"[");
            int count = 0;
            for (int index = 0; index < subtitle.Paragraphs.Count;)
            {
                Paragraph p = subtitle.Paragraphs[index];
                index++;
                if (count > 0)
                    sb.Append(',');
                sb.Append("{\"index\":");
                sb.Append(index);
                sb.Append(",\"start\":\"");
                sb.Append(p.StartTime);
                sb.Append("\",\"end\":\"");
                sb.Append(p.EndTime);
                sb.Append("\",\"horizontal\":\"");
                sb.Append(p.Horizontal);
                sb.Append("\",\"vertical\":\"");
                sb.Append(p.Vertical);
                
                if (!string.IsNullOrEmpty(p.Justification))
                {
                    sb.Append("\",\"justification\":\"");
                    sb.Append(p.Justification);
                }
                sb.Append("\",\"text\": ");
                if (!string.IsNullOrEmpty(p.Text))
                {
                    foreach (var line in p.Text.SplitToLines())
                    {
                        sb.Append("\"");
                        sb.Append(Json.EncodeJsonText(line));
                        sb.Append("\"");
                    }
                }
                sb.Append("}");
                count++;
            }
            sb.Append(']');
            return sb.ToString().Trim();
        }

        public override void LoadSubtitle(Subtitle subtitle, List<string> lines, string fileName)
        {
            _errorCount = 0;
            var sb = new StringBuilder();
            foreach (string s in lines)
                sb.Append(s);
            var allText = sb.ToString().Trim();
            if (!allText.StartsWith("[", StringComparison.Ordinal) || !allText.Contains("\"start\""))
                return;

            foreach (var line in Json.ReadObjectArray(allText))
            {
                var s = line.Trim();
                if (s.Length > 10)
                {
                    var start = Json.ReadTag(s, "start");
                    var end = Json.ReadTag(s, "end");
                    //var textLines = Json.ReadArray(s, "text");
                    var textLines = Json.ReadTag(s, "text");
                    var horizontal = Json.ReadTag(s, "horizontal");
                    var vertical = Json.ReadTag(s, "vertical");
                    var justification = Json.ReadTag(s, "justification");
                    try
                    {
                        //if (textLines.Count == 0)
                        //{
                        //    _errorCount++;
                        //}
                        //sb.Clear();
                        //foreach (var textLine in textLines)
                        //{
                        //    sb.AppendLine(Json.DecodeJsonText(textLine));
                        //}

                        sb.Clear();
                        sb.AppendLine((Json.DecodeJsonText(textLines)).Replace("\n", Environment.NewLine).Replace("<br/>", Environment.NewLine).Replace("<br/>", Environment.NewLine));
                        //subtitle.Paragraphs.Add(new Paragraph(sb.ToString().Trim(), TimeCode.ParseToMilliseconds(start), TimeCode.ParseToMilliseconds(end)));
                        subtitle.Paragraphs.Add(new Paragraph(sb.ToString().Trim(), TimeCode.ParseToMilliseconds(start), TimeCode.ParseToMilliseconds(end),horizontal.Trim(),vertical.Trim(), justification.Trim()));
                    }
                    catch (Exception)
                    {
                        _errorCount++;
                    }
                }
            }
            subtitle.Renumber();
        }

    }
}
