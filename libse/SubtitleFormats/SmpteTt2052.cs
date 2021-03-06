﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;

namespace Nikse.SubtitleEdit.Core.SubtitleFormats
{
    /// <summary>
    /// SMPTE-TT 2052
    /// </summary>
    public class SmpteTt2052 : TimedText10
    {
        public override string Extension => ".ttml";
        public new const string NameOfFormat = "SMPTE-TT 2052";
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
            XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
            XmlElement rootDeclaration = doc.DocumentElement;
            doc.InsertBefore(xmlDeclaration, rootDeclaration);

            XmlElement ttxml = doc.CreateElement("tt");

            XmlAttribute xml = doc.CreateAttribute("xml", "lang", "en");
            xml.Value = "en";
            ttxml.Attributes.Append(xml);

            XmlAttribute xmlns = doc.CreateAttribute("xmlns", "http://www.w3.org/2000/xmlns/");
            xmlns.Value = "http://www.w3.org/ns/ttml";
            ttxml.Attributes.Append(xmlns);

            XmlAttribute tts = doc.CreateAttribute("xmlns", "tts", "http://www.w3.org/2000/xmlns/");
            tts.Value = "http://www.w3.org/ns/ttml#styling";
            ttxml.Attributes.Append(tts);

            XmlAttribute ttm = doc.CreateAttribute("xmlns", "ttm", "http://www.w3.org/2000/xmlns/");
            ttm.Value = "http://www.w3.org/ns/ttml#metadata";
            ttxml.Attributes.Append(ttm);

            XmlAttribute smpte = doc.CreateAttribute("xmlns", "smpte", "http://www.w3.org/2000/xmlns/");
            smpte.Value = "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt";
            ttxml.Attributes.Append(smpte);

            XmlAttribute m608 = doc.CreateAttribute("xmlns", "m608", "http://www.w3.org/2000/xmlns/");
            m608.Value = "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt#cea608";
            ttxml.Attributes.Append(m608);

            XmlAttribute ttp = doc.CreateAttribute("xmlns", "ttp", "http://www.w3.org/2000/xmlns/");
            ttp.Value = "http://www.w3.org/ns/ttml#parameter";
            ttxml.Attributes.Append(ttp);

            XmlAttribute timeBase = doc.CreateAttribute("xmlns", "timeBase", "http://www.w3.org/2000/xmlns/");
            ttp.Value = "media";
            ttxml.Attributes.Append(ttp);

            XmlAttribute frameRate = doc.CreateAttribute("xmlns", "frameRate", "http://www.w3.org/2000/xmlns/");

            string FrameRate = "29.97";
            if (Configuration.Settings.General != null) ;
            FrameRate = Configuration.Settings.General.CurrentFrameRate.ToString();

            ttp.Value = FrameRate;
            ttxml.Attributes.Append(ttp);

            XmlAttribute frameRateMultiplier = doc.CreateAttribute("xmlns", "frameRateMultiplier", "http://www.w3.org/2000/xmlns/");
            ttp.Value = "1000 1001";
            ttxml.Attributes.Append(ttp);

            doc.AppendChild(ttxml);
            #endregion

            XmlElement xmlhead = doc.CreateElement("head"); //Creating an element with Head which is paresnt here 
            XmlElement xmlmetadata = doc.CreateElement("metadata");//Creating an element named metadata which is child to head element
            //SMPTE-TT 2052 subtitle
            XmlElement xmlttm = doc.CreateElement("ttm", "desc", "http://www.w3.org/ns/ttml#metadata");
            xmlttm.InnerXml = "SMPTE Timed Text document created by Subtitle Edit";

            XmlElement xmlinformation = doc.CreateElement("smpte", "information", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");

            XmlAttribute xmlm608 = doc.CreateAttribute("xmlns", "m608", "http://www.w3.org/2000/xmlns/");
            xmlm608.Value = "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt#cea608";
            xmlinformation.Attributes.Append(xmlm608);


            XmlAttribute xmlorigin = doc.CreateAttribute("xmlns", "origin", "http://www.w3.org/2000/xmlns/");
            xmlorigin.Value = "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt#cea608";
            xmlinformation.Attributes.Append(xmlorigin);

            XmlAttribute timemode = doc.CreateAttribute("mode");
            timemode.Value = "Preserved";
            xmlinformation.Attributes.Append(ttp);

            XmlAttribute xmlmchannel = doc.CreateAttribute("xmlns", "channel", "http://www.w3.org/2000/xmlns/");
            xmlmchannel.Value = "CC1";
            xmlinformation.Attributes.Append(xmlmchannel);

            XmlAttribute timeprogramName = doc.CreateAttribute("xmlns", "programName", "http://www.w3.org/2000/xmlns/");
            timeprogramName.Value = "Demo";
            xmlinformation.Attributes.Append(timeprogramName);

            XmlAttribute xmlcaptionService = doc.CreateAttribute("xmlns", "captionService", "http://www.w3.org/2000/xmlns/");
            xmlcaptionService.Value = "F1C1CC";
            xmlinformation.Attributes.Append(xmlcaptionService);


            xmlmetadata.AppendChild(xmlttm);
            xmlmetadata.AppendChild(xmlinformation);
            xmlhead.AppendChild(xmlmetadata);
            ttxml.AppendChild(xmlhead);

            XmlElement xmlbody = doc.CreateElement("body"); //Creating an element with body which is paresnt here 
            XmlElement xmlbiv = doc.CreateElement("div");

            xmlbody.AppendChild(xmlbiv);
            ttxml.AppendChild(xmlbody);
            return doc;
        }
        private void AddAdditionalExistingStyles(XmlDocument xml,List<XmlAttribute> attrList,ref XmlNode currentStyle)
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
            var div = xml.DocumentElement.SelectSingleNode("body", nsmgr).SelectSingleNode("div", nsmgr);
            bool hasTopCenterRegion = false;

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
                    
                    string origin,extent;

                    //horizantalPercent = 10 + Convert.ToDecimal((p.Horizontal == 0 ? 1.0 : p.Horizontal) * 3.125);
                    //verticalPercent = 10 + Convert.ToDecimal(p.Vertical * 6.66);
                    //decimal x = 10 * (1 - horizantalPercent / 100);
                    //decimal y = 10 * (1 - verticalPercent / 100);
                    var lines = text.SplitToLines();
                    int maxLineLength = lines.Max(x => x.Length);

                    double originX = 10.0 + (p.Horizontal * 2.5);
                    double originY = 10 + (p.Vertical * 5.33);
                    double extentX = (maxLineLength * 2.5);
                    double extentY =  5.33;

                    originX = originX < 85 ? originX : 85;
                    extentX = extentX < 85 ? extentX : 85;
                    origin = originX + "% " + originY + "%";
                    extent = extentX + "% "+ extentY + "%";

                    //horizantalPercent = Convert.ToDecimal((p.Horizontal == 0 ? 1.0 : p.Horizontal) * 6.66);
                    //verticalPercent = Convert.ToDecimal(p.Vertical * 3.125);
                    //origin = x+"% " + y+"%";
                    //extent = x + "% 5.33%";
                    //origin = verticalPercent + ", " + horizantalPercent + "";
                    //extent = "10% 5.33%";
                    XmlAttribute originattr = xml.CreateAttribute("tts:origin", paragraph.NamespaceURI);
                    originattr.InnerText = origin;
                    paragraph.Attributes.Append(originattr);

                    XmlAttribute extentattr = xml.CreateAttribute("tts:extent", paragraph.NamespaceURI);
                    extentattr.InnerText = extent;
                    paragraph.Attributes.Append(extentattr);
                }

                bool first = true,styleEnded=false;
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
            string xmlString = xml.InnerXml.Replace(" xmlns=\"\"", string.Empty).Replace(" xmlns:tts=\"http://www.w3.org/ns/10/ttml#style\">", ">").Replace("<br />", "<br/>").Replace("xmlns:tts=\"http://www.w3.org/ns/ttml\"", "").Replace("\n", "<br/>");

            return xmlString;
        }

    }
}