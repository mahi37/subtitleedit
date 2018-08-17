using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;

namespace Nikse.SubtitleEdit.Core.SubtitleFormats
{
    /// <summary>
    /// EBU Timed text
    /// </summary>
    public class EBUTimedText : TimedText10
    {
        public override string Extension => ".ebutt";
        public new const string NameOfFormat = "EBU Timed Text";
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
                        
            XmlElement ttxml = doc.CreateElement("tt", "tt", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
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

            XmlAttribute xmlextent = doc.CreateAttribute("tts", "extent","http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            xmlextent.Value = "704px 576px";
            ttxml.Attributes.Append(xmlextent);

            XmlAttribute frameRate = doc.CreateAttribute("ttp", "frameRate", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            string FrameRate = "29.97";
            if (Configuration.Settings.General != null)
                FrameRate = Configuration.Settings.General.CurrentFrameRate.ToString();
            frameRate.Value = FrameRate;
            ttxml.Attributes.Append(frameRate);

            XmlAttribute frameRateMultiplier = doc.CreateAttribute("ttp", "frameRateMultiplier", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            frameRateMultiplier.Value = "1 1";
            ttxml.Attributes.Append(frameRateMultiplier);

            XmlAttribute markerMode = doc.CreateAttribute("ttp", "markerMode", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            markerMode.Value = "discontinuous";
            ttxml.Attributes.Append(markerMode);

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
            #endregion

            XmlElement xmlhead = doc.CreateElement("tt:head"); //Creating an element with Head which is paresnt here 
            XmlElement xmlmetadata = doc.CreateElement("tt","metadata", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");//Creating 
            XmlElement docMetadata = doc.CreateElement("ebuttm", "documentMetadata", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            xmlmetadata.AppendChild(docMetadata);

            XmlElement docVersion = doc.CreateElement("ebuttm", "documentEbuttVersion", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            docVersion.InnerXml = "v1.0";

            XmlElement docTotSubtitiles = doc.CreateElement("ebuttm", "documentTotalNumberOfSubtitles", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            docTotSubtitiles.InnerXml = "11";
            XmlElement docMaxdispCharRow = doc.CreateElement("ebuttm", "documentMaximumNumberOfDisplayableCharacterInAnyRow", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            docMaxdispCharRow.InnerXml = "40";
            XmlElement docstartProgram = doc.CreateElement("ebuttm", "documentStartOfProgramme", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            docstartProgram.InnerXml = "00:00:00:00";
            XmlElement docOrigin = doc.CreateElement("ebuttm", "documentCountryOfOrigin", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            docOrigin.InnerXml = "DE";
            XmlElement docPublisher = doc.CreateElement("ebuttm", "documentPublisher", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            docPublisher.InnerXml = "Institut fuer Rundfunktechnik";

            docMetadata.AppendChild(docTotSubtitiles);
            docMetadata.AppendChild(docMaxdispCharRow);
            docMetadata.AppendChild(docstartProgram);
            docMetadata.AppendChild(docOrigin);
            docMetadata.AppendChild(docPublisher);

            xmlmetadata.AppendChild(docMetadata);
            xmlhead.AppendChild(xmlmetadata);


            XmlElement xmlstyling = doc.CreateElement("tt:styling");

            XmlElement xmlstyle = doc.CreateElement("tt", "style", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
             
            XmlAttribute defaultStyle = doc.CreateAttribute("xml", "id", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            defaultStyle.Value = "defaultStyle";
            xmlstyle.Attributes.Append(defaultStyle);

            XmlAttribute WhiteOnBlack = doc.CreateAttribute("tts", "fontFamily", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            WhiteOnBlack.Value = "monospaceSansSerif";
            xmlstyle.Attributes.Append(WhiteOnBlack);

            XmlAttribute fontSize = doc.CreateAttribute("tts", "fontSize", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            fontSize.Value = "1c 1c";
            xmlstyle.Attributes.Append(fontSize);

            XmlAttribute lineHeight = doc.CreateAttribute("tts", "lineHeight", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            lineHeight.Value = "normal";
            xmlstyle.Attributes.Append(lineHeight);

            XmlAttribute textAlign = doc.CreateAttribute("tt", "textAlign", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            textAlign.Value = "center";
            xmlstyle.Attributes.Append(textAlign);

            XmlAttribute color = doc.CreateAttribute("tt", "color", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            color.Value = "transparent";
            xmlstyle.Attributes.Append(color);

            XmlAttribute backgroundColor = doc.CreateAttribute("tt", "backgroundColor", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            backgroundColor.Value = "center";
            xmlstyle.Attributes.Append(backgroundColor);

            XmlAttribute fontStyle = doc.CreateAttribute("tt", "fontStyle", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            fontStyle.Value = "normal";
            xmlstyle.Attributes.Append(fontStyle);

            XmlAttribute fontWeight = doc.CreateAttribute("tt", "fontWeight", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            fontWeight.Value = "normal";
            xmlstyle.Attributes.Append(fontWeight);

            XmlAttribute textDecoration = doc.CreateAttribute("tt", "textDecoration", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            textDecoration.Value = "none";
            xmlstyle.Attributes.Append(textDecoration);

            xmlstyling.AppendChild(xmlstyle);

            XmlElement xmlstyleWhiteblack = doc.CreateElement("tt", "style", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");

            XmlAttribute stylecolor = doc.CreateAttribute("tt", "color", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            stylecolor.Value = "white";
            xmlstyleWhiteblack.Attributes.Append(stylecolor);

            XmlAttribute id = doc.CreateAttribute("tt", "id", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            id.Value = "WhiteOnBlack";
            xmlstyleWhiteblack.Attributes.Append(id);

            XmlAttribute stylebackgroundColor = doc.CreateAttribute("tt", "backgroundColor", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            stylebackgroundColor.Value = "black";
            xmlstyleWhiteblack.Attributes.Append(stylebackgroundColor);

            XmlAttribute stylefontSize = doc.CreateAttribute("tt", "fontSize", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            stylefontSize.Value = "normal";
            xmlstyleWhiteblack.Attributes.Append(stylefontSize);
            
            xmlstyling.AppendChild(xmlstyleWhiteblack);


            XmlElement xmlstyleCenter = doc.CreateElement("tt", "style", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");

            XmlAttribute styleid = doc.CreateAttribute("tt", "id", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            styleid.Value = "textCenter";
            xmlstyleCenter.Attributes.Append(styleid);

            XmlAttribute styletextAlign = doc.CreateAttribute("tt", "textAlign", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            styletextAlign.Value = "Center";
            xmlstyleCenter.Attributes.Append(styletextAlign);            

            xmlstyling.AppendChild(xmlstyleCenter);
                        
            XmlElement layout = doc.CreateElement("tt","layout", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
                                  
            XmlElement xmlalignregion = doc.CreateElement("tt", "region", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");

            XmlAttribute xmlid = doc.CreateAttribute("xml", "id", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            xmlid.Value = "top";
            xmlalignregion.Attributes.Append(xmlid);

            XmlAttribute xmlorigin = doc.CreateAttribute("tts", "origin", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            xmlorigin.Value = "10% 10%";
            xmlalignregion.Attributes.Append(xmlorigin);

            XmlAttribute xmlextents = doc.CreateAttribute("tts", "extent", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            xmlextents.Value = "80% 80%";
            xmlalignregion.Attributes.Append(xmlextents);

            XmlAttribute padding = doc.CreateAttribute("tts", "padding", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            padding.Value = "0c";
            xmlalignregion.Attributes.Append(padding);

            XmlAttribute displayAlign = doc.CreateAttribute("tts", "displayAlign", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            displayAlign.Value = "before";
            xmlalignregion.Attributes.Append(displayAlign);

            XmlAttribute writingMode = doc.CreateAttribute("tts", "writingMode", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            writingMode.Value = "lrtb";
            xmlalignregion.Attributes.Append(writingMode);

            layout.AppendChild(xmlalignregion);

            XmlElement xmlalignregionbottom = doc.CreateElement("tt", "region", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");

            XmlAttribute xmlidbottom = doc.CreateAttribute("xml", "id", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            xmlid.Value = "bottom";
            xmlalignregionbottom.Attributes.Append(xmlid);

            XmlAttribute xmloriginbottom = doc.CreateAttribute("tts", "origin", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            xmloriginbottom.Value = "10% 10%";
            xmlalignregionbottom.Attributes.Append(xmloriginbottom);

            XmlAttribute xmlextentbottom = doc.CreateAttribute("tts", "extent", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            xmlextentbottom.Value = "80% 80%";
            xmlalignregionbottom.Attributes.Append(xmlextentbottom);

            XmlAttribute paddingbottom = doc.CreateAttribute("tts", "padding", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            paddingbottom.Value = "0c";
            xmlalignregionbottom.Attributes.Append(paddingbottom);

            XmlAttribute displayAlignbottom = doc.CreateAttribute("tts", "displayAlign", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            displayAlignbottom.Value = "after";
            xmlalignregionbottom.Attributes.Append(displayAlignbottom);

            XmlAttribute writingModebottom = doc.CreateAttribute("tts", "writingMode", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            writingModebottom.Value = "lrtb";
            xmlalignregionbottom.Attributes.Append(writingModebottom);

            layout.AppendChild(xmlalignregionbottom);

            xmlhead.AppendChild(xmlmetadata);
            xmlhead.AppendChild(xmlstyling);
            xmlhead.AppendChild(layout);
            //xmlmetadata.AppendChild(xmlinformation);            
            ttxml.AppendChild(xmlhead);

            //XmlElement xmlbody = doc.CreateElement("tt:body"); //Creating an element with body which is paresnt here 
            XmlElement xmlbody = doc.CreateElement("tt", "body", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            XmlElement xmldiv = doc.CreateElement("tt", "div", "http://www.smpte-ra.org/schemas/2052-1/2010/smpte-tt");
            //XmlElement xmlbiv = doc.CreateElement("tt:div");

            xmlbody.AppendChild(xmldiv);
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
            //var div = xml.DocumentElement.SelectNodes("tt:body").SelectSingleNode("body", nsmgr).SelectSingleNode("div", nsmgr);
            var div = xml.DocumentElement.LastChild.LastChild;
            foreach (var p in subtitle.Paragraphs)
            {
                XmlNode paragraph = xml.CreateElement("tt:p", "http://www.w3.org/ns/ttml");
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
                    
                    text = text.Replace(match.Value,"");
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
            string xmlString = xml.InnerXml.Replace(" xmlns=\"\"", string.Empty).Replace(" xmlns:tts=\"http://www.w3.org/ns/10/ttml#style\">", ">").Replace("<br />", "<tt:br/>").Replace("xmlns:tts=\"http://www.w3.org/ns/ttml\"", "").Replace("\n", "<tt:br/>");

            return xmlString;
        }

    }
}