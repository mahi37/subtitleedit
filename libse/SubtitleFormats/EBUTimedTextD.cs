using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;

namespace Nikse.SubtitleEdit.Core.SubtitleFormats
{
    /// <summary>
    /// EBU Timed text D
    /// </summary>
    public class EBUTimedTextD : TimedText10
    {
        public override string Extension => ".ebuttd";
        public new const string NameOfFormat = "EBU Timed Text Distribution";
        public override string Name => NameOfFormat;
        public override bool IsMine(List<string> lines, string fileName)
        {
            if (fileName != null && !(fileName.EndsWith(Extension, StringComparison.OrdinalIgnoreCase) || fileName.EndsWith(".ttml", StringComparison.OrdinalIgnoreCase)))
                return false;

            var sb = new StringBuilder();
            lines.ForEach(line => sb.AppendLine(line));
            if (!sb.ToString().Contains("http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt#cea608"))
                return false;

            return base.IsMine(lines, fileName);
        }

        public override void LoadSubtitle(Subtitle subtitle, List<string> lines, string fileName)
        {
            var oldTimedText10TimeCodeFormat = Configuration.Settings.SubtitleSettings.TimedText10TimeCodeFormat;
            Configuration.Settings.SubtitleSettings.TimedText10TimeCodeFormat = "hh:mm:ss.ff";
            base.LoadSubtitle(subtitle, lines, fileName);
            Configuration.Settings.SubtitleSettings.TimedText10TimeCodeFormat = oldTimedText10TimeCodeFormat;
        }

        public static List<KeyValuePair<string, string>> AllignmentDictionary = new List<KeyValuePair<string, string>>
        {
            new KeyValuePair<string,string>("tts:origin='10% 10%' tts:extent='10% 5.33%'",               "{\\an7}" ),  //top left left
            new KeyValuePair<string,string>("tts:origin= '45% 10%' tts:extent= '10% 5.33%'",             "{\\an8}" ),  //top Center center	
            new KeyValuePair<string,string>("tts:origin= '80% 10%' tts:extent= '10% 5.33%'",             "{\\an9}" ),  //top right right
            new KeyValuePair<string,string>("tts:origin= '10% 47.33%' tts:extent= '10% 5.33%'",          "{\\an4}" ),  //Center left left
            new KeyValuePair<string,string>("tts:origin= '45% 47.33%' tts:extent= '10% 5.33%'",          "{\\an5}" ),  //Center Center Center
            new KeyValuePair<string,string>("tts:origin= '80% 47.33%' tts:extent= '10% 5.33%'",          "{\\an6}" ),  //Center right right           
            new KeyValuePair<string,string>("tts:origin= '10% 84.66%' tts:extent= '10% 5.33%'",          "{\\an1}" ),  //bottom left left
            new KeyValuePair<string,string>("tts:origin= '45% 84.66%' tts:extent= '10% 5.33%'",          "{\\an2}" ),  //bottom Center Center
            new KeyValuePair<string,string>("tts:origin= '80% 84.66%' tts:extent= '10% 5.33%'",          "{\\an3}" ),  //bottom right right
        };
        public string Posistioning(string position)
        {
            var code = AllignmentDictionary.FirstOrDefault(x => x.Value == position);
            if (code.Equals(new KeyValuePair<string, string>()))
                return null;
            return code.Key;
        }
        public override bool HasStyleSupport => false;

        XmlDocument declaration()
        {
            XmlDocument doc = new XmlDocument();
            #region  xml declaration 
            //(1) the xml declaration is recommended, but not mandatory
            //XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
            //XmlElement rootDeclaration = doc.DocumentElement;
            //doc.InsertBefore(xmlDeclaration, rootDeclaration);

            XmlElement ttxml = doc.CreateElement("tt", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            //XmlNode ttxml = doc.DocumentElement;
            doc.AppendChild(ttxml);

            XmlAttribute timeBase = doc.CreateAttribute("ttp", "timeBase", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            timeBase.Value = "smpte";
            ttxml.Attributes.Append(timeBase);

            XmlAttribute xml = doc.CreateAttribute("xml", "lang", "en");
            xml.Value = "en";
            ttxml.Attributes.Append(xml);

            XmlAttribute xmlcellResolution = doc.CreateAttribute("ttp", "cellResolution", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            xmlcellResolution.Value = "50 30";
            ttxml.Attributes.Append(xmlcellResolution);

            XmlAttribute xmlns = doc.CreateAttribute("xmlns", "tt", "http://www.w3.org/2000/xmlns/");
            xmlns.Value = "http://www.w3.org/ns/ttml";
            ttxml.Attributes.Append(xmlns);

            XmlAttribute xmlttp = doc.CreateAttribute("xmlns", "ttp", "http://www.w3.org/2000/xmlns/");
            xmlttp.Value = "http://www.w3.org/ns/ttml#parameter";
            ttxml.Attributes.Append(xmlttp);

            XmlAttribute xmltts = doc.CreateAttribute("xmlns", "tts", "http://www.w3.org/2000/xmlns/");
            xmltts.Value = "http://www.w3.org/ns/ttml#styling";
            ttxml.Attributes.Append(xmltts);

            

            XmlAttribute ttm = doc.CreateAttribute("xmlns", "ebuttm", "http://www.w3.org/2000/xmlns/");
            ttm.Value = "urn:ebu:tt:metadata";
            ttxml.Attributes.Append(ttm);

            //doc.AppendChild(ttxml);
            #endregion
            XmlElement xmlhead = doc.CreateElement("head");
            XmlElement xmlmetadata = doc.CreateElement("metadata");
            XmlElement docMetadata = doc.CreateElement("ebuttm","documentMetadata", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            xmlmetadata.AppendChild(docMetadata);

            XmlElement child = doc.CreateElement("ebuttm","conformsToStandard", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            child.InnerXml = "urn:ebu:tt:distribution:2014-01";
            docMetadata.AppendChild(child);

            XmlElement child2 = doc.CreateElement("ebuttm","conformsToStandard", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            child2.InnerXml = "urn:ebu:tt:distribution:2014-01";
            docMetadata.AppendChild(child);
            docMetadata.AppendChild(child2);
            xmlmetadata.AppendChild(docMetadata);
            xmlhead.AppendChild(xmlmetadata);







            //XmlElement xmlhead = doc.CreateElement("head"); //Creating an element with Head which is paresnt here 
            //            XmlElement xmlmetadata = doc.CreateElement("metadata");//Creating an element named metadata which is child to head element

            //XmlElement docMetadata = doc.CreateElement("ebuttm", "documentMetadata", "http://www.w3.org/2000/xmlns/");

            //XmlElement docVersion = doc.CreateElement("ebuttm", "conformsToStandard", "http://www.w3.org/2000/xmlns/");
            //docVersion.InnerXml = "urn:ebu:tt:distribution:2014-01";
            //XmlElement docTotSubtitiles = doc.CreateElement("ebuttm", "conformsToStandard", "http://www.w3.org/2000/xmlns/");
            //docTotSubtitiles.InnerXml = "http://www.w3.org/ns/ttml/profile/imsc1/text";
            
            //docMetadata.AppendChild(docVersion);
            //docMetadata.AppendChild(docTotSubtitiles);
            //xmlmetadata.AppendChild(docTotSubtitiles);

            XmlElement xmlstyling = doc.CreateElement("styling");

            XmlElement xmlstyle = doc.CreateElement("style", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");

            XmlAttribute defaultStyle = doc.CreateAttribute("xml", "id", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            defaultStyle.Value = "spanStyle";
            xmlstyle.Attributes.Append(defaultStyle);

            XmlAttribute WhiteOnBlack = doc.CreateAttribute("tts", "fontFamily", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            WhiteOnBlack.Value = "monospace";
            xmlstyle.Attributes.Append(WhiteOnBlack);

            XmlAttribute fontSize = doc.CreateAttribute("tts", "fontSize", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            fontSize.Value = "160%";
            xmlstyle.Attributes.Append(fontSize);

            XmlAttribute color = doc.CreateAttribute("tts", "color", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            color.Value = "#ffffff";
            xmlstyle.Attributes.Append(color);

            XmlAttribute backgroundColor = doc.CreateAttribute("tts", "backgroundColor", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            backgroundColor.Value = "#000000";
            xmlstyle.Attributes.Append(backgroundColor);

            xmlstyling.AppendChild(xmlstyle);

            XmlElement xmlstyleWhiteblack = doc.CreateElement("style", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            
            XmlAttribute id = doc.CreateAttribute("xml", "id", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            id.Value = "paragraphStyle";
            xmlstyleWhiteblack.Attributes.Append(id);

            XmlAttribute stylewrap = doc.CreateAttribute("tts", "wrapOption", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            stylewrap.Value = "noWrap";
            xmlstyleWhiteblack.Attributes.Append(stylewrap);
            
            XmlAttribute styletextAlign = doc.CreateAttribute("tts", "textAlign", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            styletextAlign.Value = "left";
            xmlstyleWhiteblack.Attributes.Append(styletextAlign);

            xmlstyling.AppendChild(xmlstyleWhiteblack);
            
            XmlElement layout = doc.CreateElement("layout");

            XmlElement xmlalignregion = doc.CreateElement("region", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");

            XmlAttribute xmlid = doc.CreateAttribute("xml", "id", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            xmlid.Value = "bottom";
            xmlalignregion.Attributes.Append(xmlid);

            XmlAttribute xmlorigin = doc.CreateAttribute("tts", "origin", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            xmlorigin.Value = "10% 10%";
            xmlalignregion.Attributes.Append(xmlorigin);

            XmlAttribute xmlextents = doc.CreateAttribute("tts", "extent", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            xmlextents.Value = "80% 80%";
            xmlalignregion.Attributes.Append(xmlextents);

            XmlAttribute displayAlign = doc.CreateAttribute("tts", "displayAlign", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            displayAlign.Value = "after";
            xmlalignregion.Attributes.Append(displayAlign);

            XmlAttribute writingMode = doc.CreateAttribute("tts", "overflow", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            writingMode.Value = "hidden";
            xmlalignregion.Attributes.Append(writingMode);

            layout.AppendChild(xmlalignregion);
            
            xmlhead.AppendChild(xmlmetadata);
            xmlhead.AppendChild(xmlstyling);
            xmlhead.AppendChild(layout);
            //xmlmetadata.AppendChild(xmlinformation);            
            ttxml.AppendChild(xmlhead);

            //XmlElement xmlbody = doc.CreateElement("tt:body"); //Creating an element with body which is paresnt here 
            XmlElement xmlbody = doc.CreateElement("body", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            XmlElement xmldiv = doc.CreateElement("div", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            //XmlElement xmlbiv = doc.CreateElement("tt:div");

            xmlbody.AppendChild(xmldiv);
            ttxml.AppendChild(xmlbody);
            //doc.AppendChild(ttxml);            
            return doc;
        }
        private void AddAdditionalExistingStyles(XmlDocument xml, List<XmlAttribute> attrList, ref XmlNode currentStyle)
        {
            XmlAttribute attr;// = xml.CreateAttribute("tts:fontStyle", "http://www.w3.org/ns/10/ttml#style");            
            foreach (XmlAttribute item in attrList) // Loop through List with foreach
            {
                attr = xml.CreateAttribute(item.Name, "http://www.w3.org/ns/10/ttml#style");
                attr.Value = item.Value;
                currentStyle.Attributes.Append(attr);
            }
        }
        public override string ToText(Subtitle subtitle, string title)
        {
            XmlDocument xml = declaration();
            var nsmgr = new XmlNamespaceManager(xml.NameTable);
            nsmgr.AddNamespace("ttml", "http://www.w3.org/ns/ttml");
            //var div = xml.DocumentElement.SelectNodes("tt:body").SelectSingleNode("body", nsmgr).SelectSingleNode("div", nsmgr);
            var div = xml.DocumentElement.LastChild.LastChild;
            foreach (var p in subtitle.Paragraphs)
            {
                XmlNode paragraph = xml.CreateElement("p", "http://www.w3.org/ns/ttml");
                string text = p.Text;

                XmlAttribute start = xml.CreateAttribute("begin");
                start.InnerText = string.Format(CultureInfo.InvariantCulture, "{0:00}:{1:00}:{2:00}:{3:00}", p.StartTime.Hours, p.StartTime.Minutes, p.StartTime.Seconds, MillisecondsToFramesMaxFrameRate(p.StartTime.Milliseconds));
                paragraph.Attributes.Append(start);

                XmlAttribute end = xml.CreateAttribute("end");
                end.InnerText = string.Format(CultureInfo.InvariantCulture, "{0:00}:{1:00}:{2:00}:{3:00}", p.EndTime.Hours, p.EndTime.Minutes, p.EndTime.Seconds, MillisecondsToFramesMaxFrameRate(p.EndTime.Milliseconds));
                paragraph.Attributes.Append(end);

                if (text.Contains("{\\an", StringComparison.Ordinal)) // this condition is used to replace 
                {
                    // if position comes after the formatting.
                    System.Text.RegularExpressions.Regex reg = new System.Text.RegularExpressions.Regex(@"{\\an\d}");
                    System.Text.RegularExpressions.Match match = reg.Match(text);

                    text = text.Replace(match.Value, "");
                    string totext = Posistioning(match.Value);
                    string[] slpittext = totext.Split('\'');

                    XmlAttribute originattr = xml.CreateAttribute("tts:origin", paragraph.NamespaceURI);
                    originattr.InnerText = slpittext[1];
                    paragraph.Attributes.Append(originattr);

                    XmlAttribute extentattr = xml.CreateAttribute("tts:extent", paragraph.NamespaceURI);
                    extentattr.InnerText = slpittext[3];
                    paragraph.Attributes.Append(extentattr);
                }
                else if (p.Vertical > 0)
                {
                   // decimal horizantalPercent, verticalPercent;
                    string origin, extent;
                    //horizantalPercent = Convert.ToDecimal((p.Horizontal == 0 ? 1.0 : p.Horizontal) * 6.66);
                    //verticalPercent = Convert.ToDecimal(p.Vertical * 3.125);


                    //decimal x = 10 * (1 - horizantalPercent / 100);
                    //decimal y = 10 * (1 - verticalPercent / 100);
                    //origin = x + "% " + y + "%";
                    //extent = x + "% 5.33%";
                    var lines = text.SplitToLines();
                    int maxLineLength = lines.Max(x => x.Length);

                    double originX = 10.0 + (p.Horizontal * 2.5);
                    double originY = 10 + (p.Vertical * 5.33);
                    double extentX = (maxLineLength * 2.5);
                    double extentY = 5.33;

                    originX = originX < 85 ? originX : 85;
                    extentX = extentX < 85 ? extentX : 85;
                    origin = originX + "% " + originY + "%";
                    extent = extentX + "% " + extentY + "%";
                    XmlAttribute originattr = xml.CreateAttribute("tts:origin", paragraph.NamespaceURI);
                    originattr.InnerText = origin;
                    paragraph.Attributes.Append(originattr);

                    XmlAttribute extentattr = xml.CreateAttribute("tts:extent", paragraph.NamespaceURI);
                    extentattr.InnerText = extent;
                    paragraph.Attributes.Append(extentattr);
                }

                bool first = true, styleEnded = false;
                bool italicOn = false;
                string line = text;

                var styles = new Stack<XmlNode>();
                XmlNode currentStyle = xml.CreateTextNode(string.Empty);
                paragraph.AppendChild(currentStyle);

                List<XmlAttribute> attrList = new List<XmlAttribute>();
                int tagsCount = 0;
                int skipCount = 0;
                for (int i = 0; i < line.Length; i++)
                {
                    if (skipCount > 0)
                    {
                        skipCount--;
                    }
                    else
                    {
                        if (line.Substring(i).StartsWith("<i>", StringComparison.Ordinal))
                        {
                            styles.Push(currentStyle);
                            currentStyle = xml.CreateNode(XmlNodeType.Element, "span", null);
                            paragraph.AppendChild(currentStyle);
                            XmlAttribute attr = xml.CreateAttribute("tts:fontStyle", "http://www.w3.org/ns/10/ttml#style");
                            attr.Value = "italic";
                            currentStyle.Attributes.Append(attr);
                            AddAdditionalExistingStyles(xml, attrList, ref currentStyle);

                            attrList.Add(attr);
                            skipCount = 2; tagsCount++;
                            italicOn = true;
                        }
                        if (line.Substring(i).StartsWith("<b>", StringComparison.Ordinal))
                        {
                            currentStyle = xml.CreateNode(XmlNodeType.Element, "span", null);
                            paragraph.AppendChild(currentStyle);
                            XmlAttribute attr = xml.CreateAttribute("tts:fontWeight", "http://www.w3.org/ns/10/ttml#style");
                            attr.Value = "bold";
                            currentStyle.Attributes.Append(attr);
                            AddAdditionalExistingStyles(xml, attrList, ref currentStyle);
                            attrList.Add(attr);
                            skipCount = 2; tagsCount++;
                            styles.Push(currentStyle);
                        }
                        if (line.Substring(i).StartsWith("<u>", StringComparison.Ordinal))
                        {
                            currentStyle = xml.CreateNode(XmlNodeType.Element, "span", null);
                            paragraph.AppendChild(currentStyle);
                            XmlAttribute attr = xml.CreateAttribute("tts:textDecoration", "http://www.w3.org/ns/10/ttml#style");
                            attr.Value = "underline";
                            currentStyle.Attributes.Append(attr);
                            AddAdditionalExistingStyles(xml, attrList, ref currentStyle);
                            attrList.Add(attr);
                            skipCount = 2; tagsCount++;
                            styles.Push(currentStyle);
                        }
                        if (line.Substring(i).StartsWith("<font ", StringComparison.Ordinal))
                        {
                            int endIndex = line.Substring(i + 1).IndexOf('>');
                            if (endIndex > 0)
                            {
                                skipCount = endIndex + 1;
                                string fontContent = line.Substring(i, skipCount);
                                if (fontContent.Contains(" color="))
                                {
                                    var arr = fontContent.Substring(fontContent.IndexOf(" color=", StringComparison.Ordinal) + 7).Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                    if (arr.Length > 0)
                                    {
                                        string fontColor = arr[0].Trim('\'').Trim('"').Trim('\'');
                                        currentStyle = xml.CreateNode(XmlNodeType.Element, "span", null);
                                        paragraph.AppendChild(currentStyle);
                                        XmlAttribute attr = xml.CreateAttribute("tts:color", "http://www.w3.org/ns/10/ttml#style");
                                        attr.Value = fontColor;
                                        currentStyle.Attributes.Append(attr);
                                        AddAdditionalExistingStyles(xml, attrList, ref currentStyle);
                                        attrList.Add(attr);
                                        styles.Push(currentStyle);
                                    }
                                }
                            }
                            else
                            {
                                skipCount = line.Length;
                            }
                            tagsCount++;
                        }
                        if (line.Substring(i).StartsWith("</b>", StringComparison.Ordinal) || line.Substring(i).StartsWith("</i>", StringComparison.Ordinal) || line.Substring(i).StartsWith("</u>", StringComparison.Ordinal) || line.Substring(i).StartsWith("</font>", StringComparison.Ordinal))
                        {
                            if (line.Substring(i).StartsWith("</b>", StringComparison.Ordinal) || line.Substring(i).StartsWith("</i>", StringComparison.Ordinal) || line.Substring(i).StartsWith("</u>", StringComparison.Ordinal) || line.Substring(i).StartsWith("</font>", StringComparison.Ordinal))
                            {
                                currentStyle = xml.CreateTextNode(string.Empty);
                                if (styles.Count > 0)
                                {
                                    currentStyle = styles.Pop().CloneNode(true);
                                    currentStyle.InnerText = string.Empty;
                                    attrList.RemoveAt(attrList.Count - 1);
                                    styleEnded = true;
                                }
                                if (line.Substring(i).StartsWith("</font>", StringComparison.Ordinal))
                                    skipCount = 6;
                                else
                                    skipCount = 3;
                                italicOn = false;

                                if (styleEnded && attrList.Count > 0)
                                {
                                    currentStyle = xml.CreateNode(XmlNodeType.Element, "span", null);
                                    AddAdditionalExistingStyles(xml, attrList, ref currentStyle);
                                    paragraph.AppendChild(currentStyle);
                                    styleEnded = false;
                                }
                            }
                            tagsCount--;
                        }
                        else if (!first)
                        {
                            if (i == 0 && italicOn && !(line.Substring(i).StartsWith("<i>", StringComparison.Ordinal)))
                            {
                                styles.Push(currentStyle);
                                currentStyle = xml.CreateNode(XmlNodeType.Element, "span", null);
                                paragraph.AppendChild(currentStyle);
                                XmlAttribute attr = xml.CreateAttribute("tts:fontStyle", "http://www.w3.org/ns/10/ttml#style");
                                attr.Value = "italic";
                                currentStyle.Attributes.Append(attr);
                                styles.Push(currentStyle);
                            }
                            currentStyle.InnerText = currentStyle.InnerText + line[i];
                        }
                        else if (first && (!(line.Substring(i).StartsWith("<", StringComparison.Ordinal))))
                            currentStyle.InnerText = currentStyle.InnerText + line[i];
                    }
                    first = false;
                    currentStyle.InnerText = currentStyle.InnerText.Replace("<", "").Replace(">", "");
                }
                div.AppendChild(paragraph);
            }
            string xmlString = xml.InnerXml.Replace(" xmlns=\"\"", string.Empty).Replace(" xmlns:tts=\"http://www.w3.org/ns/10/ttml#style\">", ">").Replace("<br />", "<tt:br/>").Replace("xmlns:tts=\"http://www.w3.org/ns/ttml\"", "").Replace("\n", "<tt:br/>");

            return xmlString;
        }

    }
}