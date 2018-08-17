using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Nikse.SubtitleEdit.Core.SubtitleFormats
{

    /// <summary>
    /// http://www.theneitherworld.com/mcpoodle/SCC_TOOLS/DOCS/SCC_FORMAT.HTML
    /// § 15.119 47 CFR Ch. I (10–1–10 Edition) (pdf)
    /// Maximum four lines + max 32 characters on each line
    /// </summary>
    public class ScenaristClosedCaptions : SubtitleFormat
    {
        public class SccPositionAndStyle
        {
            public Color ForeColor { get; set; }
            public FontStyle Style { get; set; }
            public int X { get; set; }
            public int Y { get; set; }

            public SccPositionAndStyle(Color color, FontStyle style, int y, int x)
            {
                ForeColor = color;
                Style = style;
                X = x;
                Y = y;
            }
        }

        //00:01:00:29   9420 9420 94ae 94ae 94d0 94d0 4920 f761 7320 ...    semi colon (instead of colon) before frame number is used to indicate drop frame
        private const string TimeCodeRegEx = @"^\d+:\d\d:\d\d[:,]\d\d\t";
        private static readonly Regex Regex = new Regex(TimeCodeRegEx, RegexOptions.Compiled);
        protected virtual Regex RegexTimeCodes => Regex;
        protected bool DropFrame = false;
        public static List<KeyValuePair<string, string>> AllignmentDictionary = new List<KeyValuePair<string, string>>
        {
            new KeyValuePair<string,string>("91d0",                  "{\\an7}" ),  //white top left
            new KeyValuePair<string,string>("91d6",                  "{\\an8}" ),  //trans reg top mid
            new KeyValuePair<string,string>("915e",                  "{\\an9}" ),  //trans reg top right
            //new KeyValuePair<string,string>("91dc",                  "{\\an9}" ),  //trans reg top right
            new KeyValuePair<string,string>("1670",                  "{\\an4}" ),  //white reg mid left
            new KeyValuePair<string,string>("1676",                  "{\\an5}" ),  //trans reg mid mid
            new KeyValuePair<string,string>("167c",                  "{\\an6}" ),  //trans reg mid right
            new KeyValuePair<string,string>("9470",                  "{\\an1}" ),  //white reg bot left
            new KeyValuePair<string,string>("9476",                  "{\\an2}" ),  //trans reg bot mid 
            new KeyValuePair<string,string>("947c",                  "{\\an3}" ),   //trans reg bot right
        };
        public static List<KeyValuePair<string, string>> LetterDictionary = new List<KeyValuePair<string, string>>
        {
            new KeyValuePair<string,string>("20",                  " " ),
            new KeyValuePair<string,string>("a1",                  "!" ),
            new KeyValuePair<string,string>("a2",                  "\""),
            new KeyValuePair<string,string>("23",                  "#" ),
            new KeyValuePair<string,string>("a4",                  "$" ),
            new KeyValuePair<string,string>("25",                  "%" ),
            new KeyValuePair<string,string>("26",                  "&" ),
            new KeyValuePair<string,string>("a7",                  "'" ),
            new KeyValuePair<string,string>("a8",                  "(" ),
            new KeyValuePair<string,string>("29",                  ")" ),
            new KeyValuePair<string,string>("2a",                  "á" ),
            new KeyValuePair<string,string>("ab",                  "+" ),
            new KeyValuePair<string,string>("2c",                  "," ),
            new KeyValuePair<string,string>("ad",                  "-" ),
            new KeyValuePair<string,string>("ae",                  "." ),
            new KeyValuePair<string,string>("2f",                  "/" ),
            new KeyValuePair<string,string>("b0",                  "0" ),
            new KeyValuePair<string,string>("31",                  "1" ),
            new KeyValuePair<string,string>("32",                  "2" ),
            new KeyValuePair<string,string>("b3",                  "3" ),
            new KeyValuePair<string,string>("34",                  "4" ),
            new KeyValuePair<string,string>("b5",                  "5" ),
            new KeyValuePair<string,string>("b6",                  "6" ),
            new KeyValuePair<string,string>("37",                  "7" ),
            new KeyValuePair<string,string>("38",                  "8" ),
            new KeyValuePair<string,string>("b9",                  "9" ),
            new KeyValuePair<string,string>("ba",                  ":" ),
            new KeyValuePair<string,string>("3b",                  ";" ),
            new KeyValuePair<string,string>("bc",                  "<" ),
            new KeyValuePair<string,string>("3d",                  "=" ),
            new KeyValuePair<string,string>("3e",                  ">" ),
            new KeyValuePair<string,string>("bf",                  "?" ),
            new KeyValuePair<string,string>("40",                  "@" ),
            new KeyValuePair<string,string>("c1",                  "A" ),
            new KeyValuePair<string,string>("c2",                  "B" ),
            new KeyValuePair<string,string>("43",                  "C" ),
            new KeyValuePair<string,string>("c4",                  "D" ),
            new KeyValuePair<string,string>("45",                  "E" ),
            new KeyValuePair<string,string>("46",                  "F" ),
            new KeyValuePair<string,string>("c7",                  "G" ),
            new KeyValuePair<string,string>("c8",                  "H" ),
            new KeyValuePair<string,string>("49",                  "I" ),
            new KeyValuePair<string,string>("4a",                  "J" ),
            new KeyValuePair<string,string>("cb",                  "K" ),
            new KeyValuePair<string,string>("4c",                  "L" ),
            new KeyValuePair<string,string>("cd",                  "M" ),
            new KeyValuePair<string,string>("ce",                  "N" ),
            new KeyValuePair<string,string>("4f",                  "O" ),
            new KeyValuePair<string,string>("d0",                  "P" ),
            new KeyValuePair<string,string>("51",                  "Q" ),
            new KeyValuePair<string,string>("52",                  "R" ),
            new KeyValuePair<string,string>("d3",                  "S" ),
            new KeyValuePair<string,string>("54",                  "T" ),
            new KeyValuePair<string,string>("d5",                  "U" ),
            new KeyValuePair<string,string>("d6",                  "V" ),
            new KeyValuePair<string,string>("57",                  "W" ),
            new KeyValuePair<string,string>("58",                  "X" ),
            new KeyValuePair<string,string>("d9",                  "Y" ),
            new KeyValuePair<string,string>("da",                  "Z" ),
            new KeyValuePair<string,string>("5b",                  "[" ),
            new KeyValuePair<string,string>("dc",                  "é" ),
            new KeyValuePair<string,string>("5d",                  "]" ),
            new KeyValuePair<string,string>("5e",                  "í" ),
            new KeyValuePair<string,string>("df",                  "ó" ),
            new KeyValuePair<string,string>("e0",                  "ú" ),
            new KeyValuePair<string,string>("61",                  "a" ),
            new KeyValuePair<string,string>("62",                  "b" ),
            new KeyValuePair<string,string>("e3",                  "c" ),
            new KeyValuePair<string,string>("64",                  "d" ),
            new KeyValuePair<string,string>("e5",                  "e" ),
            new KeyValuePair<string,string>("e6",                  "f" ),
            new KeyValuePair<string,string>("67",                  "g" ),
            new KeyValuePair<string,string>("68",                  "h" ),
            new KeyValuePair<string,string>("e9",                  "i" ),
            new KeyValuePair<string,string>("ea",                  "j" ),
            new KeyValuePair<string,string>("6b",                  "k" ),
            new KeyValuePair<string,string>("ec",                  "l" ),
            new KeyValuePair<string,string>("6d",                  "m" ),
            new KeyValuePair<string,string>("6e",                  "n" ),
            new KeyValuePair<string,string>("ef",                  "o" ),
            new KeyValuePair<string,string>("70",                  "p" ),
            new KeyValuePair<string,string>("f1",                  "q" ),
            new KeyValuePair<string,string>("f2",                  "r" ),
            new KeyValuePair<string,string>("73",                  "s" ),
            new KeyValuePair<string,string>("f4",                  "t" ),
            new KeyValuePair<string,string>("75",                  "u" ),
            new KeyValuePair<string,string>("76",                  "v" ),
            new KeyValuePair<string,string>("f7",                  "w" ),
            new KeyValuePair<string,string>("f8",                  "x" ),
            new KeyValuePair<string,string>("fb",                  "ç" ),
            new KeyValuePair<string,string>("79",                  "y" ),
            new KeyValuePair<string,string>("7a",                  "z" ),
            new KeyValuePair<string,string>("fb",                  ""  ),
            new KeyValuePair<string,string>("7c",                  ""  ),
            new KeyValuePair<string,string>("fd",                  "Ñ" ),
            new KeyValuePair<string,string>("fe",                  "ñ" ),
            new KeyValuePair<string,string>("7f",                  "■" ),
            new KeyValuePair<string,string>("7b",                  "ç" ),
            new KeyValuePair<string,string>("63",                  "c" ),
            new KeyValuePair<string,string>("65",                  "e" ),
            new KeyValuePair<string,string>("66",                  "f" ),
            new KeyValuePair<string,string>("69",                  "i" ),
            new KeyValuePair<string,string>("6a",                  "j" ),
            new KeyValuePair<string,string>("6c",                  "l" ),
            new KeyValuePair<string,string>("6e",                  "n" ),
            new KeyValuePair<string,string>("6f",                  "o" ),
            new KeyValuePair<string,string>("71",                  "q" ),
            new KeyValuePair<string,string>("72",                  "r" ),
            new KeyValuePair<string,string>("74",                  "t" ),
            new KeyValuePair<string,string>("77",                  "w" ),
            new KeyValuePair<string,string>("78",                  "x" ),
            new KeyValuePair<string,string>("91b0",                ""  ),
            new KeyValuePair<string,string>("9131",                "°" ),
            new KeyValuePair<string,string>("9132",                "½" ),
            new KeyValuePair<string,string>("91b3",                ""  ),
            new KeyValuePair<string,string>("9134",                ""  ),
            new KeyValuePair<string,string>("91b5",                ""  ),
            new KeyValuePair<string,string>("91b6",                "£" ),
            new KeyValuePair<string,string>("9137",                "♪" ),
            new KeyValuePair<string,string>("9138",                "à" ),
            new KeyValuePair<string,string>("91b9",                ""  ),
            new KeyValuePair<string,string>("91ba",                "è" ),
            new KeyValuePair<string,string>("913b",                "â" ),
            new KeyValuePair<string,string>("91bc",                "ê" ),
            new KeyValuePair<string,string>("913d",                "î" ),
            new KeyValuePair<string,string>("913e",                "ô" ),
            new KeyValuePair<string,string>("91bf",                "û" ),
            new KeyValuePair<string,string>("9220",                "Á" ),
            new KeyValuePair<string,string>("92a1",                "É" ),
            new KeyValuePair<string,string>("92a2",                "Ó" ),
            new KeyValuePair<string,string>("9223",                "Ú" ),
            new KeyValuePair<string,string>("92a4",                "Ü" ),
            new KeyValuePair<string,string>("9225",                "ü" ),
            new KeyValuePair<string,string>("9226",                "'" ),
            new KeyValuePair<string,string>("92a7",                "i" ),
            new KeyValuePair<string,string>("92a8",                "*" ),
            new KeyValuePair<string,string>("9229",                "'" ),
            new KeyValuePair<string,string>("922a",                "-" ),
            new KeyValuePair<string,string>("92ab",                ""  ),
            new KeyValuePair<string,string>("922c",                ""  ),
            new KeyValuePair<string,string>("92ad",                "\""),
            new KeyValuePair<string,string>("92ae",                "\""),
            new KeyValuePair<string,string>("922f",                ""  ),
            new KeyValuePair<string,string>("92b0",                "À" ),
            new KeyValuePair<string,string>("9231",                "Â" ),
            new KeyValuePair<string,string>("9232",                ""  ),
            new KeyValuePair<string,string>("92b3",                "È" ),
            new KeyValuePair<string,string>("9234",                "Ê" ),
            new KeyValuePair<string,string>("92b5",                "Ë" ),
            new KeyValuePair<string,string>("92b6",                "ë" ),
            new KeyValuePair<string,string>("9237",                "Î" ),
            new KeyValuePair<string,string>("9238",                "Ï" ),
            new KeyValuePair<string,string>("92b9",                "ï" ),
            new KeyValuePair<string,string>("92b3",                "ô" ),
            new KeyValuePair<string,string>("923b",                "Ù" ),
            new KeyValuePair<string,string>("92b3",                "ù" ),
            new KeyValuePair<string,string>("923d",                "Û" ),
            new KeyValuePair<string,string>("923e",                ""  ),
            new KeyValuePair<string,string>("92bf",                ""  ),
            new KeyValuePair<string,string>("1320",                "Ã" ),
            new KeyValuePair<string,string>("c1 1320",             "Ã" ),
            new KeyValuePair<string,string>("c180 1320",           "Ã" ),
            new KeyValuePair<string,string>("13a1",                "ã" ),
            new KeyValuePair<string,string>("80 13a1",             "ã" ),
            new KeyValuePair<string,string>("6180 13a1",           "ã" ),
            new KeyValuePair<string,string>("13a2",                "Í" ),
            new KeyValuePair<string,string>("49 13a2",             "Í" ),
            new KeyValuePair<string,string>("4980 13a2",           "Í" ),
            new KeyValuePair<string,string>("1323",                "Ì" ),
            new KeyValuePair<string,string>("49 1323",             "Ì" ),
            new KeyValuePair<string,string>("4980 1323",           "Ì" ),
            new KeyValuePair<string,string>("13a4",                "ì" ),
            new KeyValuePair<string,string>("e9 13a4",             "ì" ),
            new KeyValuePair<string,string>("e980 13a4",           "ì" ),
            new KeyValuePair<string,string>("1325",                "Ò" ),
            new KeyValuePair<string,string>("4f 1325",             "Ò" ),
            new KeyValuePair<string,string>("4f80 1325",           "Ò" ),
            new KeyValuePair<string,string>("1326",                "ò" ),
            new KeyValuePair<string,string>("ef 1326",             "ò" ),
            new KeyValuePair<string,string>("ef80 1326",           "ò" ),
            new KeyValuePair<string,string>("13a7",                "Õ" ),
            new KeyValuePair<string,string>("4f 13a7",             "Õ" ),
            new KeyValuePair<string,string>("4f80 13a7",           "Õ" ),
            new KeyValuePair<string,string>("13a8",                "õ" ),
            new KeyValuePair<string,string>("ef 13a8",             "õ" ),
            new KeyValuePair<string,string>("ef80 13a8",           "õ" ),
            new KeyValuePair<string,string>("1329",                "{" ),
            new KeyValuePair<string,string>("132a",                "}" ),
            new KeyValuePair<string,string>("13ab",                "\\"),
            new KeyValuePair<string,string>("132c",                "^" ),
            new KeyValuePair<string,string>("13ad",                "_" ),
            new KeyValuePair<string,string>("13ae",                "|" ),
            new KeyValuePair<string,string>("132f",                "~" ),
            new KeyValuePair<string,string>("13b0",                "Ä" ),
            new KeyValuePair<string,string>("c180 13b0",           "Ä" ),
            new KeyValuePair<string,string>("1331",                "ä" ),
            new KeyValuePair<string,string>("6180 1331",           "ä" ),
            new KeyValuePair<string,string>("1332",                "Ö" ),
            new KeyValuePair<string,string>("4f80 1332",           "Ö" ),
            new KeyValuePair<string,string>("13b3",                "ö" ),
            new KeyValuePair<string,string>("ef80 13b3",           "ö" ),
            new KeyValuePair<string,string>("1334",                ""  ),
            new KeyValuePair<string,string>("13b5",                ""  ),
            new KeyValuePair<string,string>("13b6",                ""  ),
            new KeyValuePair<string,string>("1337",                "|" ),
            new KeyValuePair<string,string>("1338",                "Å" ),
            new KeyValuePair<string,string>("13b9",                "å" ),
            new KeyValuePair<string,string>("13b3",                "Ø" ),
            new KeyValuePair<string,string>("133b",                "ø" ),
            new KeyValuePair<string,string>("13b3",                ""  ),
            new KeyValuePair<string,string>("133d",                ""  ),
            new KeyValuePair<string,string>("133e",                ""  ),
            new KeyValuePair<string,string>("13bf",                ""  ),
            new KeyValuePair<string,string>("9420",                ""  ), //9420=RCL, Resume Caption Loadin
            new KeyValuePair<string,string>("94ae",                ""  ), //94ae=Clear Buffer
            new KeyValuePair<string,string>("942c",                ""  ), //942c=Clear Caption
            new KeyValuePair<string,string>("8080",                ""  ), //8080=Wait One Frame
            new KeyValuePair<string,string>("942f",                ""  ), //942f=Display Caption
            new KeyValuePair<string,string>("9440",                ""  ), //9440=? first sub?
            new KeyValuePair<string,string>("9452",                ""  ), //?
            new KeyValuePair<string,string>("9454",                ""  ), //?
            new KeyValuePair<string,string>("9470",                ""  ), //9470=?
            new KeyValuePair<string,string>("94d0",                ""  ), //94d0=?
            new KeyValuePair<string,string>("94d6",                ""  ), //94d6=?
            new KeyValuePair<string,string>("942f",                ""  ), //942f=End of Caption
            new KeyValuePair<string,string>("94f2",                ""  ),
            new KeyValuePair<string,string>("94f4",                ""  ),
            new KeyValuePair<string,string>("9723",                " " ), // ?
            new KeyValuePair<string,string>("97a1",                " " ), // ?
            new KeyValuePair<string,string>("97a2",                " " ), // ?
            new KeyValuePair<string,string>("1370",                ""  ), //1370=?
            new KeyValuePair<string,string>("13e0",                ""  ), //13e0=?
            new KeyValuePair<string,string>("13f2",                ""  ), //13f2=?
            new KeyValuePair<string,string>("136e",                ""  ), //136e=?
            new KeyValuePair<string,string>("94ce",                ""  ), //94ce=?
            new KeyValuePair<string,string>("2c2f",                ""  ), //?
            new KeyValuePair<string,string>("1130",                "®" ),
            new KeyValuePair<string,string>("1131",                "°" ),
            new KeyValuePair<string,string>("1132",                "½" ),
            new KeyValuePair<string,string>("1133",                "¿" ),
            new KeyValuePair<string,string>("1134",                "TM"),
            new KeyValuePair<string,string>("1135",                "¢" ),
            new KeyValuePair<string,string>("1136",                "£" ),
            new KeyValuePair<string,string>("1137",                "♪" ),
            new KeyValuePair<string,string>("1138",                "à" ),
            new KeyValuePair<string,string>("1138",                " " ), // transparent space
            new KeyValuePair<string,string>("113a",                "è" ),
            new KeyValuePair<string,string>("113b",                "â" ),
            new KeyValuePair<string,string>("113c",                "ê" ),
            new KeyValuePair<string,string>("113d",                "î" ),
            new KeyValuePair<string,string>("113e",                "ô" ),
            new KeyValuePair<string,string>("113f",                "û" ),
            new KeyValuePair<string,string>("9130",                "®" ),
            new KeyValuePair<string,string>("9131",                "°" ),
            new KeyValuePair<string,string>("9132",                "½" ),
            new KeyValuePair<string,string>("9133",                "¿" ),
            new KeyValuePair<string,string>("9134",                "TM"),
            new KeyValuePair<string,string>("9135",                "¢" ),
            new KeyValuePair<string,string>("9136",                "£" ),
            new KeyValuePair<string,string>("9137",                "♪" ),
            new KeyValuePair<string,string>("9138",                "à" ),
            new KeyValuePair<string,string>("9138",                " " ), // transparent space
            new KeyValuePair<string,string>("913a",                "è" ),
            new KeyValuePair<string,string>("913b",                "â" ),
            new KeyValuePair<string,string>("913c",                "ê" ),
            new KeyValuePair<string,string>("913d",                "î" ),
            new KeyValuePair<string,string>("913e",                "ô" ),
            new KeyValuePair<string,string>("913f",                "û" ),
            new KeyValuePair<string,string>("a180 92a7 92a7",      "¡" ),
            new KeyValuePair<string,string>("92a7 92a7",           "¡" ),
            new KeyValuePair<string,string>("91b3 91b3",           "¿" ),

            new KeyValuePair<string,string>("6180 9138 9138",      "à"), //61=a
            new KeyValuePair<string,string>("9138 9138",           "à"),

            new KeyValuePair<string,string>("6180 913b 913b",      "â"),
            new KeyValuePair<string,string>("913b 913b",           "â"),

            new KeyValuePair<string,string>("6180 1331 1331",      "ä"),
            new KeyValuePair<string,string>("1331 1331",           "ä"),

            new KeyValuePair<string,string>("e580 91ba 91ba",      "è"),
            new KeyValuePair<string,string>("6180 91ba 91ba",      "è"),
            new KeyValuePair<string,string>("91ba 91ba",           "è"),

            new KeyValuePair<string,string>("e580 91bc 91bc",      "ê"),
            new KeyValuePair<string,string>("6180 91bc 91bc",      "ê"),
            new KeyValuePair<string,string>("91bc 91bc",           "ê"),


            new KeyValuePair<string,string>("e580 92b6 92b6",      "ë"), //e5=e (+65?)
            new KeyValuePair<string,string>("6580 92b6 92b6",      "ë"),
            new KeyValuePair<string,string>("92b6 92b6",           "ë"),

            new KeyValuePair<string,string>("e980 13a4 13a4",      "ì"), //e9 = i
            new KeyValuePair<string,string>("13a4 13a4",           "ì"),

            new KeyValuePair<string,string>("e980 913d 913d",      "î"),
            new KeyValuePair<string,string>("913d 913d",           "î"),

            new KeyValuePair<string,string>("e980 92b9 92b9",      "ï"),
            new KeyValuePair<string,string>("92b9 92b9",           "ï"),

            new KeyValuePair<string,string>("1326 1326",           "ò"), //o=ef or 6f
            new KeyValuePair<string,string>("ef80 1326 1326",      "ò"),
            new KeyValuePair<string,string>("6f80 1326 1326",      "ò"),

            new KeyValuePair<string,string>("913e 913e",           "ô"),
            new KeyValuePair<string,string>("ef80 913e 913e",      "ô"),
            new KeyValuePair<string,string>("6f80 913e 913e",      "ô"),

            new KeyValuePair<string,string>("13b3 13b3",           "ö"),
            new KeyValuePair<string,string>("ef80 13b3 13b3",      "ö"),
            new KeyValuePair<string,string>("6f80 13b3 13b3",      "ö"),

            new KeyValuePair<string,string>("7580 13b3 13b3",      "ù"), //u=75
            new KeyValuePair<string,string>("13b3 13b3",           "ù"),

            new KeyValuePair<string,string>("7580 92bc 92bc",      "ù"),
            new KeyValuePair<string,string>("92bc 92bc",           "ù"),

            new KeyValuePair<string,string>("7580 91bf 91bf",      "û"),
            new KeyValuePair<string,string>("91bf 91bf",           "û"),

            new KeyValuePair<string,string>("7580 9225 9225",      "ü"),
            new KeyValuePair<string,string>("9225 9225",           "ü"),

            new KeyValuePair<string,string>("4380 9232 9232",      "Ç"), //43=C
            new KeyValuePair<string,string>("9232 9232",           "Ç"),

            new KeyValuePair<string,string>("c180 1338 1338",      "Å"), //c1=A
            new KeyValuePair<string,string>("1338 1338",           "Å"),

            new KeyValuePair<string,string>("c180 1338 1338",      "Å"),
            new KeyValuePair<string,string>("1338 1338",           "Å"),

            new KeyValuePair<string,string>("c180 92b0 92b0",      "À"),
            new KeyValuePair<string,string>("92b0 92b0",           "À"),

            new KeyValuePair<string,string>("c180 9220 9220",      "Á"),
            new KeyValuePair<string,string>("9220 9220",           "Á"),

            new KeyValuePair<string,string>("c180 9231 9231",      "Â"),
            new KeyValuePair<string,string>("9231 9231",           "Â"),

            new KeyValuePair<string,string>("c180 1320 1320",      "Ã"),
            new KeyValuePair<string,string>("1320 1320",           "Ã"),

            new KeyValuePair<string,string>("c180 13b0 13b0",      "Ä"),
            new KeyValuePair<string,string>("13b0 13b0",           "Ä"),

            new KeyValuePair<string,string>("c180 1320 1320",      "Ã"),
            new KeyValuePair<string,string>("1320 1320",           "Ã"),

            new KeyValuePair<string,string>("c180 13b0 13b0",      "Ä"),
            new KeyValuePair<string,string>("13b0 13b0",           "Ä"),

            new KeyValuePair<string,string>("4580 92b3 92b3",      "È"),
            new KeyValuePair<string,string>("92b3 92b3",           "È"),

            new KeyValuePair<string,string>("4580 92a1 92a1",      "É"),
            new KeyValuePair<string,string>("92a1 92a1",           "É"),

            new KeyValuePair<string,string>("4580 9234 9234",      "Ê"),
            new KeyValuePair<string,string>("9234 9234",           "Ê"),

            new KeyValuePair<string,string>("4580 92b5 92b5",      "Ë"),
            new KeyValuePair<string,string>("92b5 92b5",           "Ë"),

            new KeyValuePair<string,string>("4980 1323 1323",      "Ì"),
            new KeyValuePair<string,string>("1323 1323",           "Ì"),

            new KeyValuePair<string,string>("4980 13a2 13a2",      "Í"),
            new KeyValuePair<string,string>("13a2 13a2",           "Í"),

            new KeyValuePair<string,string>("4980 9237 9237",      "Î"),
            new KeyValuePair<string,string>("9237 9237",           "Î"),

            new KeyValuePair<string,string>("4980 9238 9238",      "Ï"),
            new KeyValuePair<string,string>("9238 9238",           "Ï"),

            new KeyValuePair<string,string>("4f80 92a2 92a2",      "Ó"), //4f=O
            new KeyValuePair<string,string>("92a2 92a2",           "Ó"),

            new KeyValuePair<string,string>("4f80 1325 1325",      "Ò"),
            new KeyValuePair<string,string>("1325 1325",           "Ò"),

            new KeyValuePair<string,string>("4f80 92ba 92ba",      "Ô"),
            new KeyValuePair<string,string>("92ba 92ba",           "Ô"),

            new KeyValuePair<string,string>("4f80 13a7 13a7",      "Õ"),
            new KeyValuePair<string,string>("13a7 13a7",           "Õ"),

            new KeyValuePair<string,string>("4f80 1332 1332",      "Ö"),
            new KeyValuePair<string,string>("1332 1332",           "Ö"),

            new KeyValuePair<string,string>("d580 923b 923b",      "Ù"), //d5=U
            new KeyValuePair<string,string>("923b 923b",           "Ù"),

            new KeyValuePair<string,string>("d580 9223 9223",      "Ú"),
            new KeyValuePair<string,string>("923d 923d",           "Û"),

            new KeyValuePair<string,string>("d580 923b 923b",      "Ù"),
            new KeyValuePair<string,string>("9223 9223",           "Ú"),

            new KeyValuePair<string,string>("d580 92a4 92a4",      "Ü"),
            new KeyValuePair<string,string>("92a4 92a4",           "Ü"),

            new KeyValuePair<string,string>("d580 923d 923d",      "Û"),
            new KeyValuePair<string,string>("923d 923d",           "Û"),
        };
        // list of position related codes
        private static List<KeyValuePair<string, SccPositionAndStyle>> PositionDictionary = new List<KeyValuePair<string, SccPositionAndStyle>>
        {
             new KeyValuePair<string,SccPositionAndStyle>("91d0", new SccPositionAndStyle(Color.White, FontStyle.Regular, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1140", new SccPositionAndStyle(Color.White, FontStyle.Regular, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1160", new SccPositionAndStyle(Color.White, FontStyle.Regular, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1240", new SccPositionAndStyle(Color.White, FontStyle.Regular, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1260", new SccPositionAndStyle(Color.White, FontStyle.Regular, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1540", new SccPositionAndStyle(Color.White, FontStyle.Regular, 5, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1560", new SccPositionAndStyle(Color.White, FontStyle.Regular, 6, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1640", new SccPositionAndStyle(Color.White, FontStyle.Regular, 7, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1660", new SccPositionAndStyle(Color.White, FontStyle.Regular, 8, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1740", new SccPositionAndStyle(Color.White, FontStyle.Regular, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1760", new SccPositionAndStyle(Color.White, FontStyle.Regular, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1040", new SccPositionAndStyle(Color.White, FontStyle.Regular, 11, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1340", new SccPositionAndStyle(Color.White, FontStyle.Regular, 12, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1360", new SccPositionAndStyle(Color.White, FontStyle.Regular, 13, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1440", new SccPositionAndStyle(Color.White, FontStyle.Regular, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1460", new SccPositionAndStyle(Color.White, FontStyle.Regular, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1141", new SccPositionAndStyle(Color.White, FontStyle.Underline, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1161", new SccPositionAndStyle(Color.White, FontStyle.Underline, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1241", new SccPositionAndStyle(Color.White, FontStyle.Underline, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1261", new SccPositionAndStyle(Color.White, FontStyle.Underline, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1541", new SccPositionAndStyle(Color.White, FontStyle.Underline, 5, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1561", new SccPositionAndStyle(Color.White, FontStyle.Underline, 6, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1641", new SccPositionAndStyle(Color.White, FontStyle.Underline, 7, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1661", new SccPositionAndStyle(Color.White, FontStyle.Underline, 8, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1741", new SccPositionAndStyle(Color.White, FontStyle.Underline, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1761", new SccPositionAndStyle(Color.White, FontStyle.Underline, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1041", new SccPositionAndStyle(Color.White, FontStyle.Underline, 11, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1341", new SccPositionAndStyle(Color.White, FontStyle.Underline, 12, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1361", new SccPositionAndStyle(Color.White, FontStyle.Underline, 13, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1441", new SccPositionAndStyle(Color.White, FontStyle.Underline, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1461", new SccPositionAndStyle(Color.White, FontStyle.Underline, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1142", new SccPositionAndStyle(Color.Green, FontStyle.Regular, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1162", new SccPositionAndStyle(Color.Green, FontStyle.Regular, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1242", new SccPositionAndStyle(Color.Green, FontStyle.Regular, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1262", new SccPositionAndStyle(Color.Green, FontStyle.Regular, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1542", new SccPositionAndStyle(Color.Green, FontStyle.Regular, 5, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1562", new SccPositionAndStyle(Color.Green, FontStyle.Regular, 6, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1642", new SccPositionAndStyle(Color.Green, FontStyle.Regular, 7, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1662", new SccPositionAndStyle(Color.Green, FontStyle.Regular, 8, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1742", new SccPositionAndStyle(Color.Green, FontStyle.Regular, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1762", new SccPositionAndStyle(Color.Green, FontStyle.Regular, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1042", new SccPositionAndStyle(Color.Green, FontStyle.Regular, 11, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1342", new SccPositionAndStyle(Color.Green, FontStyle.Regular, 12, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1362", new SccPositionAndStyle(Color.Green, FontStyle.Regular, 13, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1442", new SccPositionAndStyle(Color.Green, FontStyle.Regular, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1462", new SccPositionAndStyle(Color.Green, FontStyle.Regular, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1143", new SccPositionAndStyle(Color.Green, FontStyle.Underline, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1163", new SccPositionAndStyle(Color.Green, FontStyle.Underline, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1243", new SccPositionAndStyle(Color.Green, FontStyle.Underline, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1263", new SccPositionAndStyle(Color.Green, FontStyle.Underline, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1543", new SccPositionAndStyle(Color.Green, FontStyle.Underline, 5, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1563", new SccPositionAndStyle(Color.Green, FontStyle.Underline, 6, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1643", new SccPositionAndStyle(Color.Green, FontStyle.Underline, 7, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1663", new SccPositionAndStyle(Color.Green, FontStyle.Underline, 8, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1743", new SccPositionAndStyle(Color.Green, FontStyle.Underline, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1763", new SccPositionAndStyle(Color.Green, FontStyle.Underline, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1043", new SccPositionAndStyle(Color.Green, FontStyle.Underline, 11, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1343", new SccPositionAndStyle(Color.Green, FontStyle.Underline, 12, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1363", new SccPositionAndStyle(Color.Green, FontStyle.Underline, 13, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1443", new SccPositionAndStyle(Color.Green, FontStyle.Underline, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1463", new SccPositionAndStyle(Color.Green, FontStyle.Underline, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1144", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1164", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1244", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1264", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1544", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 5, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1564", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 6, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1644", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 7, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1664", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 8, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1744", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1764", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1044", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 11, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1344", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 12, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1364", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 13, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1444", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1464", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1145", new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1165", new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1245", new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1265", new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1545", new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 5, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1565", new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 6, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1645", new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 7, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1665", new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 8, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1745", new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1765", new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1045", new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 11, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1345", new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 12, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1365", new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 13, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1445", new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1465", new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1146", new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1166", new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1246", new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1266", new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1546", new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 5, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1566", new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 6, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1646", new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 7, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1666", new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 8, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1746", new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1766", new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1046", new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 11, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1346", new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 12, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1366", new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 13, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1446", new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1466", new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1147", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1167", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1247", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1267", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1547", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 5, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1567", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 6, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1647", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 7, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1667", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 8, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1747", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1767", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1047", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 11, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1347", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 12, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1367", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 13, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1447", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1467", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1148", new SccPositionAndStyle(Color.Red, FontStyle.Regular, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1168", new SccPositionAndStyle(Color.Red, FontStyle.Regular, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1248", new SccPositionAndStyle(Color.Red, FontStyle.Regular, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1268", new SccPositionAndStyle(Color.Red, FontStyle.Regular, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1548", new SccPositionAndStyle(Color.Red, FontStyle.Regular, 5, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1568", new SccPositionAndStyle(Color.Red, FontStyle.Regular, 6, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1648", new SccPositionAndStyle(Color.Red, FontStyle.Regular, 7, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1668", new SccPositionAndStyle(Color.Red, FontStyle.Regular, 8, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1748", new SccPositionAndStyle(Color.Red, FontStyle.Regular, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1768", new SccPositionAndStyle(Color.Red, FontStyle.Regular, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1048", new SccPositionAndStyle(Color.Red, FontStyle.Regular, 11, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1348", new SccPositionAndStyle(Color.Red, FontStyle.Regular, 12, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1368", new SccPositionAndStyle(Color.Red, FontStyle.Regular, 13, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1448", new SccPositionAndStyle(Color.Red, FontStyle.Regular, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1468", new SccPositionAndStyle(Color.Red, FontStyle.Regular, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1149", new SccPositionAndStyle(Color.Red, FontStyle.Underline, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1169", new SccPositionAndStyle(Color.Red, FontStyle.Underline, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1249", new SccPositionAndStyle(Color.Red, FontStyle.Underline, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1269", new SccPositionAndStyle(Color.Red, FontStyle.Underline, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1549", new SccPositionAndStyle(Color.Red, FontStyle.Underline, 5, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1569", new SccPositionAndStyle(Color.Red, FontStyle.Underline, 6, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1649", new SccPositionAndStyle(Color.Red, FontStyle.Underline, 7, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1669", new SccPositionAndStyle(Color.Red, FontStyle.Underline, 8, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1749", new SccPositionAndStyle(Color.Red, FontStyle.Underline, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1769", new SccPositionAndStyle(Color.Red, FontStyle.Underline, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1049", new SccPositionAndStyle(Color.Red, FontStyle.Underline, 11, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1349", new SccPositionAndStyle(Color.Red, FontStyle.Underline, 12, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1369", new SccPositionAndStyle(Color.Red, FontStyle.Underline, 13, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1449", new SccPositionAndStyle(Color.Red, FontStyle.Underline, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1469", new SccPositionAndStyle(Color.Red, FontStyle.Underline, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("114a", new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("116a", new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("124a", new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("126a", new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("154a", new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 5, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("156a", new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 6, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("164a", new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 7, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("166a", new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 8, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("174a", new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("176a", new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("104a", new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 11, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("134a", new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 12, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("136a", new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 13, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("144a", new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("146a", new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("114b", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("116b", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("124b", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("126b", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("154b", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 5, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("156b", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 6, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("164b", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 7, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("166b", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 8, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("174b", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("176b", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("104b", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 11, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("134b", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 12, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("136b", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 13, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("144b", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("146b", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("114c", new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("116c", new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("124c", new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("126c", new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("154c", new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 5, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("156c", new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 6, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("164c", new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 7, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("166c", new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 8, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("174c", new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("176c", new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("104c", new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 11, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("134c", new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 12, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("136c", new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 13, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("144c", new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("146c", new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("114d", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("116d", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("124d", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("126d", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("154d", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 5, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("156d", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 6, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("164d", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 7, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("166d", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 8, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("174d", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("176d", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("104d", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 11, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("134d", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 12, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("136d", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 13, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("144d", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("146d", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("114e", new SccPositionAndStyle(Color.White, FontStyle.Italic, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("116e", new SccPositionAndStyle(Color.White, FontStyle.Italic, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("124e", new SccPositionAndStyle(Color.White, FontStyle.Italic, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("126e", new SccPositionAndStyle(Color.White, FontStyle.Italic, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("154e", new SccPositionAndStyle(Color.White, FontStyle.Italic, 5, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("156e", new SccPositionAndStyle(Color.White, FontStyle.Italic, 6, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("164e", new SccPositionAndStyle(Color.White, FontStyle.Italic, 7, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("166e", new SccPositionAndStyle(Color.White, FontStyle.Italic, 8, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("174e", new SccPositionAndStyle(Color.White, FontStyle.Italic, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("176e", new SccPositionAndStyle(Color.White, FontStyle.Italic, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("104e", new SccPositionAndStyle(Color.White, FontStyle.Italic, 11, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("134e", new SccPositionAndStyle(Color.White, FontStyle.Italic, 12, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("136e", new SccPositionAndStyle(Color.White, FontStyle.Italic, 13, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("144e", new SccPositionAndStyle(Color.White, FontStyle.Italic, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("146e", new SccPositionAndStyle(Color.White, FontStyle.Italic, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("114f", new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("116f", new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("124f", new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("126f", new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("154f", new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 5, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("156f", new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 6, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("164f", new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 7, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("166f", new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 8, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("174f", new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("176f", new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("104f", new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 11, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("134f", new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 12, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("136f", new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 13, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("144f", new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("146f", new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("91d0", new SccPositionAndStyle(Color.White, FontStyle.Regular, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9151", new SccPositionAndStyle(Color.White, FontStyle.Underline, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("91c2", new SccPositionAndStyle(Color.Green, FontStyle.Regular, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9143", new SccPositionAndStyle(Color.Green, FontStyle.Underline, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("91c4", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9145", new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9146", new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("91c7", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("91c8", new SccPositionAndStyle(Color.Red, FontStyle.Regular, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9149", new SccPositionAndStyle(Color.Red, FontStyle.Underline, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("914a", new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("91cb", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("914c", new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("91cd", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 1, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9170", new SccPositionAndStyle(Color.White, FontStyle.Regular, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("91f1", new SccPositionAndStyle(Color.White, FontStyle.Underline, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9162", new SccPositionAndStyle(Color.Green, FontStyle.Regular, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("91e3", new SccPositionAndStyle(Color.Green, FontStyle.Underline, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9164", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("91e5", new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("91e6", new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9167", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9168", new SccPositionAndStyle(Color.Red, FontStyle.Regular, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("91e9", new SccPositionAndStyle(Color.Red, FontStyle.Underline, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("91ea", new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("916b", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("91ec", new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("916d", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 2, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("92d0", new SccPositionAndStyle(Color.White, FontStyle.Regular, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9251", new SccPositionAndStyle(Color.White, FontStyle.Underline, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("92c2", new SccPositionAndStyle(Color.Green, FontStyle.Regular, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9243", new SccPositionAndStyle(Color.Green, FontStyle.Underline, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("92c4", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9245", new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9246", new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("92c7", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("92c8", new SccPositionAndStyle(Color.Red, FontStyle.Regular, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9249", new SccPositionAndStyle(Color.Red, FontStyle.Underline, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("924a", new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("92cb", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("924c", new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("92cd", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 3, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9270", new SccPositionAndStyle(Color.White, FontStyle.Regular, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("92f1", new SccPositionAndStyle(Color.White, FontStyle.Underline, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9262", new SccPositionAndStyle(Color.Green, FontStyle.Regular, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("92e3", new SccPositionAndStyle(Color.Green, FontStyle.Underline, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9264", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("92e5", new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("92e6", new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9267", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9268", new SccPositionAndStyle(Color.Red, FontStyle.Regular, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("92e9", new SccPositionAndStyle(Color.Red, FontStyle.Underline, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("92ea", new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("926b", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("92ec", new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("926d", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 4, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("15d0", new SccPositionAndStyle(Color.White, FontStyle.Regular, 5, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1551", new SccPositionAndStyle(Color.White, FontStyle.Underline, 5, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("15c2", new SccPositionAndStyle(Color.Green, FontStyle.Regular, 5, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("15c4", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 5, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("15c7", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 5, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("15c8", new SccPositionAndStyle(Color.Red, FontStyle.Regular, 5, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("15cb", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 5, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("15cd", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 5, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1570", new SccPositionAndStyle(Color.White, FontStyle.Regular, 6, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("15f1", new SccPositionAndStyle(Color.White, FontStyle.Underline, 6, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("15e5", new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 6, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("15e6", new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 6, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("15e9", new SccPositionAndStyle(Color.Red, FontStyle.Underline, 6, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("15ea", new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 6, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("15ec", new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 6, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("16d0", new SccPositionAndStyle(Color.White, FontStyle.Regular, 7, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1651", new SccPositionAndStyle(Color.White, FontStyle.Underline, 7, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("16c2", new SccPositionAndStyle(Color.Green, FontStyle.Regular, 7, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("16c4", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 7, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("16c7", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 7, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("16c8", new SccPositionAndStyle(Color.Red, FontStyle.Regular, 7, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("16cb", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 7, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("16cd", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 7, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1670", new SccPositionAndStyle(Color.White, FontStyle.Regular, 8, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("16f1", new SccPositionAndStyle(Color.White, FontStyle.Underline, 8, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("16e3", new SccPositionAndStyle(Color.Green, FontStyle.Underline, 8, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("16e5", new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 8, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("16e6", new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 8, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("16e9", new SccPositionAndStyle(Color.Red, FontStyle.Underline, 8, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("16ea", new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 8, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("16ec", new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 8, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("97d0", new SccPositionAndStyle(Color.White, FontStyle.Regular, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9751", new SccPositionAndStyle(Color.White, FontStyle.Underline, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("97c2", new SccPositionAndStyle(Color.Green, FontStyle.Regular, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9743", new SccPositionAndStyle(Color.Green, FontStyle.Underline, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("97c4", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9745", new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9746", new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("97c7", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("97c8", new SccPositionAndStyle(Color.Red, FontStyle.Regular, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9749", new SccPositionAndStyle(Color.Red, FontStyle.Underline, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("974a", new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("97cb", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("974c", new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("97cd", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 9, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9770", new SccPositionAndStyle(Color.White, FontStyle.Regular, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("97f1", new SccPositionAndStyle(Color.White, FontStyle.Underline, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9762", new SccPositionAndStyle(Color.Green, FontStyle.Regular, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("97e3", new SccPositionAndStyle(Color.Green, FontStyle.Underline, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9764", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("97e5", new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("97e6", new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9767", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9768", new SccPositionAndStyle(Color.Red, FontStyle.Regular, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("97e9", new SccPositionAndStyle(Color.Red, FontStyle.Underline, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("97ea", new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("976b", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("97ec", new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("976d", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 10, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("10d0", new SccPositionAndStyle(Color.White, FontStyle.Regular, 11, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1051", new SccPositionAndStyle(Color.White, FontStyle.Underline, 11, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("10c2", new SccPositionAndStyle(Color.Green, FontStyle.Regular, 11, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("10c4", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 11, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("10c7", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 11, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("10c8", new SccPositionAndStyle(Color.Red, FontStyle.Regular, 11, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("10cb", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 11, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("10cd", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 11, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("13d0", new SccPositionAndStyle(Color.White, FontStyle.Regular, 12, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1351", new SccPositionAndStyle(Color.White, FontStyle.Underline, 12, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("13c2", new SccPositionAndStyle(Color.Green, FontStyle.Regular, 12, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("13c4", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 12, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("13c7", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 12, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("13c8", new SccPositionAndStyle(Color.Red, FontStyle.Regular, 12, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("13cb", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 12, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("13cd", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 12, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("1370", new SccPositionAndStyle(Color.White, FontStyle.Regular, 13, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("13f1", new SccPositionAndStyle(Color.White, FontStyle.Underline, 13, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("13e3", new SccPositionAndStyle(Color.Green, FontStyle.Underline, 13, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("13e5", new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 13, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("13e6", new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 13, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("13e9", new SccPositionAndStyle(Color.Red, FontStyle.Underline, 13, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("13ea", new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 13, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("13ec", new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 13, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("94d0", new SccPositionAndStyle(Color.White, FontStyle.Regular, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9451", new SccPositionAndStyle(Color.White, FontStyle.Underline, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("94c2", new SccPositionAndStyle(Color.Green, FontStyle.Regular, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9443", new SccPositionAndStyle(Color.Green, FontStyle.Underline, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("94c4", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9445", new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9446", new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("94c7", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("94c8", new SccPositionAndStyle(Color.Red, FontStyle.Regular, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9449", new SccPositionAndStyle(Color.Red, FontStyle.Underline, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("944a", new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("94cb", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("944c", new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("94cd", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 14, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9470", new SccPositionAndStyle(Color.White, FontStyle.Regular, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("94f1", new SccPositionAndStyle(Color.White, FontStyle.Underline, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9462", new SccPositionAndStyle(Color.Green, FontStyle.Regular, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("94e3", new SccPositionAndStyle(Color.Green, FontStyle.Underline, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9464", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("94e5", new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("94e6", new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9467", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9468", new SccPositionAndStyle(Color.Red, FontStyle.Regular, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("94e9", new SccPositionAndStyle(Color.Red, FontStyle.Underline, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("94ea", new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("946b", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("94ec", new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("946d", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 15, 0)),
             new KeyValuePair<string,SccPositionAndStyle>("9152", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 1, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("91d3", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 1, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("9154", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 1, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("91d5", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 1, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("91d6", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 1, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("9157", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 1, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("9158", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 1, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("91d9", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 1, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("91da", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 1, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("915b", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 1, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("91dc", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 1, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("915d", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 1, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("915e", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 1, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("91df", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 1, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("91f2", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 2, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("9173", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 2, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("91f4", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 2, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("9175", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 2, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("9176", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 2, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("91f7", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 2, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("91f8", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 2, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("9179", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 2, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("917a", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 2, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("91fb", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 2, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("917c", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 2, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("91fd", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 2, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("91fe", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 2, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("917f", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 2, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("9252", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 3, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("92d3", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 3, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("9254", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 3, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("92d5", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 3, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("92d6", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 3, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("9257", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 3, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("9258", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 3, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("92d9", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 3, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("92da", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 3, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("925b", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 3, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("92dc", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 3, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("925d", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 3, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("925e", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 3, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("92df", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 3, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("92f2", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 4, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("9273", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 4, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("92f4", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 4, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("9275", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 4, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("9276", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 4, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("92f7", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 4, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("92f8", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 4, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("9279", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 4, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("927a", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 4, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("92fb", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 4, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("927c", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 4, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("92fd", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 4, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("92fe", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 4, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("927f", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 4, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("1552", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 5, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("15d3", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 5, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("1554", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 5, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("15d5", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 5, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("15d6", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 5, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("1557", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 5, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("1558", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 5, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("15d9", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 5, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("15da", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 5, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("155b", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 5, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("15dc", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 5, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("155d", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 5, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("155e", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 5, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("15df", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 5, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("15f2", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 6, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("1573", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 6, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("15f4", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 6, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("1575", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 6, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("1576", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 6, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("15f7", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 6, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("15f8", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 6, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("1579", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 6, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("157a", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 6, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("15fb", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 6, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("157c", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 6, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("15fd", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 6, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("15fe", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 6, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("157f", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 6, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("1652", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 7, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("16d3", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 7, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("1654", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 7, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("16d5", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 7, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("16d6", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 7, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("1657", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 7, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("1658", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 7, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("16d9", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 7, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("16da", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 7, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("165b", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 7, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("16dc", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 7, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("165d", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 7, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("165e", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 7, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("16df", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 7, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("16f2", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 8, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("1673", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 8, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("16f4", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 8, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("1675", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 8, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("1676", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 8, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("16f7", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 8, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("16f8", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 8, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("1679", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 8, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("167a", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 8, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("16fb", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 8, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("167c", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 8, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("16fd", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 8, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("16fe", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 8, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("167f", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 8, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("9752", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 9, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("97d3", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 9, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("9754", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 9, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("97d5", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 9, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("97d6", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 9, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("9757", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 9, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("9758", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 9, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("97d9", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 9, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("97da", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 9, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("975b", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 9, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("97dc", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 9, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("975d", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 9, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("975e", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 9, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("97df", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 9, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("97f2", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 10, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("9773", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 10, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("97f4", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 10, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("9775", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 10, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("9776", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 10, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("97f7", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 10, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("97f8", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 10, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("9779", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 10, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("977a", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 10, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("97fb", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 10, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("977c", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 10, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("97fd", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 10, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("97fe", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 10, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("977f", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 10, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("1052", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 11, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("10d3", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 11, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("1054", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 11, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("10d5", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 11, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("10d6", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 11, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("1057", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 11, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("1058", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 11, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("10d9", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 11, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("10da", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 11, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("105b", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 11, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("10dc", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 11, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("105d", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 11, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("105e", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 11, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("10df", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 11, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("1352", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 12, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("13d3", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 12, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("1354", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 12, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("13d5", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 12, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("13d6", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 12, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("1357", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 12, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("1358", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 12, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("13d9", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 12, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("13da", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 12, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("135b", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 12, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("13dc", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 12, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("135d", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 12, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("135e", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 12, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("13df", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 12, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("13f2", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 13, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("1373", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 13, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("13f4", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 13, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("1375", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 13, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("1376", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 13, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("13f7", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 13, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("13f8", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 13, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("1379", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 13, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("137a", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 13, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("13fb", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 13, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("137c", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 13, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("13fd", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 13, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("13fe", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 13, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("137f", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 13, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("9452", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 14, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("94d3", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 14, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("9454", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 14, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("94d5", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 14, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("94d6", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 14, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("9457", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 14, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("9458", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 14, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("94d9", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 14, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("94da", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 14, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("945b", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 14, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("94dc", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 14, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("945d", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 14, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("945e", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 14, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("94df", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 14, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("94f2", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 15, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("9473", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 15, 4)),
             new KeyValuePair<string,SccPositionAndStyle>("94f4", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 15, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("9475", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 15, 8)),
             new KeyValuePair<string,SccPositionAndStyle>("9476", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 15, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("94f7", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 15, 12)),
             new KeyValuePair<string,SccPositionAndStyle>("94f8", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 15, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("9479", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 15, 16)),
             new KeyValuePair<string,SccPositionAndStyle>("947a", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 15, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("94fb", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 15, 20)),
             new KeyValuePair<string,SccPositionAndStyle>("947c", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 15, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("94fd", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 15, 24)),
             new KeyValuePair<string,SccPositionAndStyle>("94fe", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 15, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("947f", new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 15, 28)),
             new KeyValuePair<string,SccPositionAndStyle>("9120", new SccPositionAndStyle(Color.White, FontStyle.Regular, -1, -1)),                                    // mid-row commands
             new KeyValuePair<string,SccPositionAndStyle>("91a1", new SccPositionAndStyle(Color.White, FontStyle.Underline, -1, -1)),
             new KeyValuePair<string,SccPositionAndStyle>("91a2", new SccPositionAndStyle(Color.Green, FontStyle.Regular, -1, -1)),
             new KeyValuePair<string,SccPositionAndStyle>("9123", new SccPositionAndStyle(Color.Green, FontStyle.Underline, -1, -1)),
             new KeyValuePair<string,SccPositionAndStyle>("91a4", new SccPositionAndStyle(Color.Blue, FontStyle.Regular, -1, -1)),
             new KeyValuePair<string,SccPositionAndStyle>("9125", new SccPositionAndStyle(Color.Blue, FontStyle.Underline, -1, -1)),
             new KeyValuePair<string,SccPositionAndStyle>("9126", new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, -1, -1)),
             new KeyValuePair<string,SccPositionAndStyle>("91a7", new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, -1, -1)),
             new KeyValuePair<string,SccPositionAndStyle>("91a8", new SccPositionAndStyle(Color.Red, FontStyle.Regular, -1, -1)),
             new KeyValuePair<string,SccPositionAndStyle>("9129", new SccPositionAndStyle(Color.Red, FontStyle.Underline, -1, -1)),
             new KeyValuePair<string,SccPositionAndStyle>("912a", new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, -1, -1)),
             new KeyValuePair<string,SccPositionAndStyle>("91ab", new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, -1, -1)),
             new KeyValuePair<string,SccPositionAndStyle>("912c", new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, -1, -1)),
             new KeyValuePair<string,SccPositionAndStyle>("91ad", new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, -1, -1)),
             new KeyValuePair<string,SccPositionAndStyle>("91ae", new SccPositionAndStyle(Color.Transparent, FontStyle.Italic, -1, -1)),
             new KeyValuePair<string,SccPositionAndStyle>("912f", new SccPositionAndStyle(Color.Transparent, FontStyle.Italic | FontStyle.Underline, -1, -1)),
             new KeyValuePair<string,SccPositionAndStyle>("94a8", new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, -1, -1)),
        };
        public override string Extension => ".scc";

        public override string Name => "Scenarist Closed Captions";

        private static string FixMax4LinesAndMax32CharsPerLine(string text, string language)
        {
            // fix attempt 1
            var lines = text.Trim().SplitToLines();
            if (IsAllOkay(lines))
                return text;

            // fix attempt 2
            text = Utilities.AutoBreakLine(text, 1, 4, language);
            lines = text.Trim().SplitToLines();
            if (IsAllOkay(lines))
                return text;

            // fix attempt 3
            text = AutoBreakLineMax4Lines(text, 32);
            lines = text.Trim().SplitToLines();
            if (IsAllOkay(lines))
                return text;

            var sb = new StringBuilder();
            int count = 0;
            foreach (string line in lines)
            {
                if (count < 4)
                {
                    if (line.Length > 32)
                        sb.AppendLine(line.Substring(0, 32));
                    else
                        sb.AppendLine(line);
                }
                count++;
            }
            return sb.ToString().Trim();
        }

        private static bool IsAllOkay(List<string> lines)
        {
            if (lines.Count > 4)
                return false;
            for (int i = 0; i < lines.Count; i++)
            {
                if (lines[i].Length > 32)
                    return false;
            }
            return true;
        }

        private static int GetLastIndexOfSpace(string s, int endCount)
        {
            var end = Math.Min(endCount, s.Length - 1);
            while (end > 0)
            {
                if (s[end] == ' ')
                    return end;
                end--;
            }
            return -1;
        }

        private static string AutoBreakLineMax4Lines(string text, int maxLength)
        {
            string s = text.Replace(Environment.NewLine, " ");
            s = s.Replace("  ", " ");
            var sb = new StringBuilder();
            int i = GetLastIndexOfSpace(s, maxLength);
            if (i > 0)
            {
                sb.AppendLine(s.Substring(0, i));
                s = s.Remove(0, i).Trim();
                if (s.Length <= maxLength)
                    i = s.Length;
                else
                    i = GetLastIndexOfSpace(s, maxLength);
                if (i > 0)
                {
                    sb.AppendLine(s.Substring(0, i));
                    s = s.Remove(0, i).Trim();
                    if (s.Length <= maxLength)
                        i = s.Length;
                    else
                        i = GetLastIndexOfSpace(s, maxLength);
                    if (i > 0)
                    {
                        sb.AppendLine(s.Substring(0, i));
                        s = s.Remove(0, i).Trim();
                        if (s.Length <= maxLength)
                            i = s.Length;
                        else
                            i = GetLastIndexOfSpace(s, maxLength);
                        if (i > 0)
                        {
                            sb.AppendLine(s.Substring(0, i));
                        }
                        else
                        {
                            sb.Append(s);
                        }
                    }
                    else
                    {
                        sb.Append(s);
                    }
                }
                return sb.ToString().Trim();
            }
            return text;
        }
        public string addSEPositions(Paragraph para)
        {
            string text = para.Text;
            if (para.Horizontal <= 11 && para.Vertical <= 5)
                text = "{\\an7}"+ text;
            else if (para.Horizontal <= 21 && para.Vertical <= 5)
                text = "{\\an8}" + text;
            else if (para.Horizontal <= 32 && para.Vertical <= 5)
                text = "{\\an9}" + text;
            else if (para.Horizontal <= 11 && para.Vertical <= 10)
                text = "{\\an4}" + text;
            else if (para.Horizontal <= 21 && para.Vertical <= 10)
                text = "{\\an5}" + text;
            else if (para.Horizontal <= 32 && para.Vertical <= 10)
                text = "{\\an6}" + text;
            else if (para.Horizontal <= 11 && para.Vertical <= 15)
                text = "{\\an1}" + text;
            else if (para.Horizontal <= 21 && para.Vertical <= 15)
                text = "{\\an2}" + text;
            else
                text = "{\\an3}" + text;

            return text;

        }
        public override string ToText(Subtitle subtitle, string title)
        {
            try
            {
                var sb = new StringBuilder();
                sb.AppendLine("Scenarist_SCC V1.0");
                sb.AppendLine();
                string language = LanguageAutoDetect.AutoDetectGoogleLanguage(subtitle);
                for (int i = 0; i < subtitle.Paragraphs.Count; i++)
                {
                    Paragraph p = subtitle.Paragraphs[i];
                    Paragraph next = subtitle.GetParagraphOrDefault(i + 1);
                    if (next != null)
                        p.EndTime.TotalMilliseconds = p.EndTime.TotalMilliseconds;

                    string timeCode = ToTimeCodebyText(p.Text, p.StartTime.TotalMilliseconds, p.EndTime.TotalMilliseconds);
                     //sb.AppendLine(string.Format("{0}\t94ae 94ae 9420 9420 {1} 942c 942c 8080 8080 942f 942f", timeCode, ToSccAllignedText(p.Text, language)));
                    sb.AppendLine(string.Format("{0}\t94ae 94ae 9420 9420 {1} 942c 942c 8080 8080 942f 942f", timeCode, ToSccAllignedText(addSEPositions(p), language)));
                    //sb.AppendLine(string.Format("{0}\t94ae 94ae 9420 9420 {1} 942c 942c 8080 8080 942f 942f", timeCode, ToSccAllignedText(p, language)));
                    sb.AppendLine();

                    if (next != null)
                    {
                        if (p.EndTime.TotalMilliseconds != next.StartTime.TotalMilliseconds)
                        {
                            double clearCaptionCode = p.EndTime.TotalMilliseconds;
                            double nextlineTime = GetTimeCodebyText(next.Text, next.StartTime.TotalMilliseconds);
                            if (nextlineTime > clearCaptionCode)
                                sb.AppendLine(string.Format("{0}\t942c 942c", ToTimeCode(clearCaptionCode)));
                        }
                    }
                    else
                        sb.AppendLine(string.Format("{0}\t942c t942c", ToTimeCode(p.EndTime.TotalMilliseconds + Math.Abs(p.EndTime.TotalMilliseconds - p.StartTime.TotalMilliseconds))));
                }
                return sb.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private string replaceAllignedChars(string rawText)
        {
            var sb = new StringBuilder();
            sb.Append(GetCodeFromAllignWord(rawText));
            return sb.ToString().Trim();
        }
        /// <summary>
        /// change the text based on subtitile alignment to scc allignments for all applications
        /// </summary>
        /// <param name="text"></param>
        /// <param name="language"></param>
        /// <returns></returns>
        private string ToSccAllignedText(string text, string language)
        {
            var sb = new StringBuilder();
            string brkCode = string.Empty;
            // if position not been added manually adding position to bottom center.
            if (!text.StartsWith("{\\an"))
                text = "{\\an2}" + text;

            if (text.Contains("{\\an1}") || text.Contains("{\\an2}") || text.Contains("{\\an3}") || text.Contains("{\\an4}") || text.Contains("{\\an5}") || text.Contains("{\\an6}") || text.Contains("{\\an7}") || text.Contains("{\\an8}") || text.Contains("{\\an9}"))
            {

                text = text.Replace("{\\an1}", "{\\an1}|").Replace("{\\an2}", "{\\an2}|").Replace("{\\an3}", "{\\an3}|").Replace("{\\an4}", "{\\an4}|").Replace("{\\an5}", "{\\an5}|").Replace("{\\an6}", "{\\an6}|").Replace("{\\an7}", "{\\an7}|").Replace("{\\an8}", "{\\an8}|").Replace("{\\an9}", "{\\an9}|");
                string[] rawSplitedText = text.Split('|');
                string posCode = replaceAllignedChars(rawSplitedText[0].Trim()) + " ";
                //center alignment
                SccPositionAndStyle sccPosition = GetColorAndPosition(posCode.Trim());

                var rawSplitText = (FixMax4LinesAndMax32CharsPerLine(rawSplitedText[1].Trim(), language)).SplitToLines();

                //get position value
                string position;
                if (text.Contains("{\\an1}") || text.Contains("{\\an4}") || text.Contains("{\\an7}"))
                    position = "Left";
                else if (text.Contains("{\\an3}") || text.Contains("{\\an6}") || text.Contains("{\\an9}"))
                    position = "Right";
                else
                    position = "Center";
                // get y axis max line value
                int yaxisvalue = 0;
                if (rawSplitedText[0] == "{\\an9}" || rawSplitedText[0] == "{\\an8}" || rawSplitedText[0] == "{\\an7}")
                    yaxisvalue = rawSplitText.Count();
                else if (rawSplitedText[0] == "{\\an6}" || rawSplitedText[0] == "{\\an5}" || rawSplitedText[0] == "{\\an4}")
                {
                    if (rawSplitedText[0] == "{\\an5}")
                        yaxisvalue = 8;
                    else
                        yaxisvalue = 8 - (rawSplitText.Count() > 2 ? rawSplitText.Count() % 2 : 0);
                }
                else
                    yaxisvalue = 15;

                
                int MaxXLength = rawSplitText.Max(x => x.Length);

                sb.Append(ToSccText(rawSplitedText[1].Trim(), language, MaxXLength, yaxisvalue - 1, position) + " ");
            }

            return sb.ToString().Trim();
        }
        private string ToSccAllignedText(Paragraph paragraph, string language)
        {
            var sb = new StringBuilder();
            string position = string.Empty;
            if (paragraph.Justification == "L")
                position = "Left";
            else if (paragraph.Justification == "R")
                position = "Right";
            else
                position = "Center";

            //sb.Append(ToSccText(paragraph.Text, language, paragraph.Horizontal, paragraph.Vertical, position) + " ");

            var rawSplitText = paragraph.Text.Trim().SplitToLines();

            int yaxisValue = paragraph.Vertical;
            if (yaxisValue == 15)
                yaxisValue = yaxisValue - rawSplitText.Count() - 1;



            sb.Append(ToSccText(paragraph.Text, language, paragraph.Horizontal, yaxisValue, position) + " ");




            return sb.ToString().Trim();
        }
        private static string ToSccText(string text, string language, int xPosition = 32, int yposition = 14, string position = "Center")
        {
            text = FixMax4LinesAndMax32CharsPerLine(text, language);
            var lines = text.Trim().SplitToLines();
            int italic = 0;
            var sb = new StringBuilder();
            int count = 1;
            foreach (string line in lines)
            {
                text = line.Trim();
                if (count > 0)
                    sb.Append(' ');
                sb.Append(GetCenterCodesDefault(text, count, lines.Count, xPosition, yposition, position));
                //sb.Append(GetCenterCodes(text, count, lines.Count, xPosition, yposition));
                count++;
                int i = 0;
                string code = string.Empty;
                if (italic > 0)
                {
                    sb.Append("91ae 91ae "); // italic
                }
                while (i < text.Length)
                {
                    string s = text.Substring(i, 1);
                    string codeFromLetter = GetCodeFromLetter(s);
                    string newCode;
                    if (text.Substring(i).StartsWith("<i>", StringComparison.Ordinal))
                    {
                        newCode = "91ae";
                        i += 2;
                        italic++;
                    }
                    else if (text.Substring(i).StartsWith("</i>", StringComparison.Ordinal) && italic > 0)
                    {
                        newCode = "9120";
                        i += 3;
                        italic--;
                    }
                    else if (text[i] == '’')
                    {
                        if (code.Length == 4)
                        {
                            sb.Append(code + " ");
                            code = string.Empty;
                        }
                        if (code.Length == 0)
                        {
                            code = "80";
                        }
                        if (code.Length == 2)
                        {
                            code += "a7";
                            sb.Append(code + " ");
                        }
                        code = "9229";
                        newCode = "";
                    }
                    else if (codeFromLetter == null)
                        newCode = GetCodeFromLetter(" ");
                    else
                        newCode = codeFromLetter;

                    if (code.Length == 2 && newCode.Length == 4)
                    {
                        code += "80";
                    }

                    if (code.Length == 4)
                    {
                        sb.Append(code + " ");
                        if (code.StartsWith('9') || code.StartsWith('8')) // control codes must be double
                            sb.Append(code + " ");
                        code = string.Empty;
                    }

                    if (code.Length == 2 && newCode.Length == 2)
                    {
                        code += newCode;
                        newCode = string.Empty;
                    }

                    if (newCode.Length == 4 && code.Length == 0)
                    {
                        code = newCode;
                    }
                    else if (newCode.Length == 2 && code.Length == 0)
                    {
                        code = newCode;
                    }
                    else if (newCode.Length > 4)
                    {
                        if (code.Length == 2)
                        {
                            code += "80";
                            sb.Append(code + " ");
                            if (code.StartsWith('9') || code.StartsWith('8')) // control codes must be double
                                sb.Append(code + " ");
                            code = string.Empty;
                        }
                        else if (code.Length == 4)
                        {
                            sb.Append(code + " ");
                            if (code.StartsWith('9') || code.StartsWith('8')) // control codes must be double
                                sb.Append(code + " ");
                            code = string.Empty;
                        }
                        sb.Append(newCode.TrimEnd() + " ");
                    }

                    i++;
                }
                if (code.Length == 2)
                    code += "80";
                if (code.Length == 4)
                    sb.Append(code);
            }

            return sb.ToString().Trim();
        }
        // get the allignemnt code based on subtitile app related position.
        private static string GetCodeFromAllignWord(string letter)
        {
            var code = AllignmentDictionary.FirstOrDefault(x => x.Value == letter);
            if (code.Equals(new KeyValuePair<string, string>()))
                return null;
            return code.Key;
        }
        private static string GetCodeFromLetter(string letter)
        {
            var code = LetterDictionary.FirstOrDefault(x => x.Value == letter);
            if (code.Equals(new KeyValuePair<string, string>()))
                return null;
            return code.Key;
        }

        private static string GetLetterFromCode(string hexCode)
        {
            var letter = LetterDictionary.FirstOrDefault(x => x.Key == hexCode.ToLowerInvariant());
            if (letter.Equals(new KeyValuePair<string, string>()))
                return null;
            return letter.Value;
        }
        public static string GetCenterCodesDefault(string text, int lineNumber, int totalLines, int MaxXposition, int MaxYposition, string position = "Center")
        {
            double left, Maxleft; int column = 0, columnRest = 0;
            int row = MaxYposition - (totalLines - lineNumber);

            var rowCodes = new List<string> { "91", "91", "92", "92", "15", "15", "16", "16", "97", "97", "10", "13", "13", "94", "94" };
            string rowCode = rowCodes[row];
            if (position == "Center")
            {
                left = (32.0 - text.Length) / 2.0;
                columnRest = (int)left % 4;
                column = (int)left - columnRest;
            }
            else if (position == "Right")
            {
                left = (32 - text.Length) / 2.0;
                Maxleft = (32 - MaxXposition) / 2.0;
                column = (int)(left + Maxleft);
                columnRest = column % 4;
                column = column - columnRest;
            }
            else
            {
                if (rowCode == "91" && totalLines > 1)
                {
                    row = lineNumber - 1 < 0 ? 0 : lineNumber - 1;
                    rowCode = rowCodes[row];
                }
                left = (MaxXposition - text.Length) / 2;
                columnRest = (int)left % 4;
                column = (int)left - columnRest;

            }


            List<string> columnCodes;
            switch (column)
            {
                case 0:
                    columnCodes = new List<string> { "d0", "70", "d0", "70", "d0", "70", "d0", "70", "d0", "70", "d0", "d0", "70", "d0", "70" };
                    break;
                case 4:
                    columnCodes = new List<string> { "52", "f2", "52", "f2", "52", "f2", "52", "f2", "52", "f2", "52", "52", "f2", "52", "f2" };
                    break;
                case 8:
                    columnCodes = new List<string> { "54", "f4", "54", "f4", "54", "f4", "54", "f4", "54", "f4", "54", "54", "f4", "54", "f4" };
                    break;
                case 12:
                    columnCodes = new List<string> { "d6", "76", "d6", "76", "d6", "76", "d6", "76", "d6", "76", "d6", "d6", "76", "d6", "76" };
                    break;
                case 16:
                    columnCodes = new List<string> { "58", "f8", "58", "f8", "58", "f8", "58", "f8", "58", "f8", "58", "58", "f8", "58", "f8" };
                    break;
                case 20:
                    columnCodes = new List<string> { "da", "7a", "da", "7a", "da", "7a", "da", "7a", "da", "7a", "da", "da", "7a", "da", "7a" };
                    break;
                case 24:
                    columnCodes = new List<string> { "dc", "7c", "dc", "7c", "dc", "7c", "dc", "7c", "dc", "7c", "dc", "dc", "7c", "dc", "7c" };
                    break;
                default: // 28
                    columnCodes = new List<string> { "5e", "fe", "5e", "fe", "5e", "fe", "5e", "fe", "5e", "fe", "5e", "5e", "fe", "5e", "fe" };
                    break;
            }
            string code = rowCode + columnCodes[row];

            if (columnRest == 1)
                return code + " " + code + " 97a1 97a1 ";
            if (columnRest == 2)
                return code + " " + code + " 97a2 97a2 ";
            if (columnRest == 3)
                return code + " " + code + " 9723 9723 ";
            return code + " " + code + " ";
        }
        public static string GetCenterCodes(string text, int lineNumber, int totalLines, int xposition, int yposition)
        {
            int row = 0;

            row = yposition - (totalLines - lineNumber);
            //row = (lineNumber - 1);
            row = row < 0 ? 0 : row;
            var rowCodes = new List<string> { "91", "91", "92", "92", "15", "15", "16", "16", "97", "97", "10", "13", "13", "94", "94" };
            string rowCode = rowCodes[row];
            if (rowCode == "91" && totalLines > 1)
            {
                row = lineNumber - 1 < 0 ? 0 : lineNumber - 1;
                rowCode = rowCodes[row];
            }
            int left = (xposition - text.Length) / 2;
            int columnRest = left % 4;
            int column = left - columnRest;


            List<string> columnCodes;
            switch (column)
            {
                case 0:
                    columnCodes = new List<string> { "d0", "70", "d0", "70", "d0", "70", "d0", "70", "d0", "70", "d0", "d0", "70", "d0", "70" };
                    break;
                case 4:
                    columnCodes = new List<string> { "52", "f2", "52", "f2", "52", "f2", "52", "f2", "52", "f2", "52", "52", "f2", "52", "f2" };
                    break;
                case 8:
                    columnCodes = new List<string> { "54", "f4", "54", "f4", "54", "f4", "54", "f4", "54", "f4", "54", "54", "f4", "54", "f4" };
                    break;
                case 12:
                    columnCodes = new List<string> { "d6", "76", "d6", "76", "d6", "76", "d6", "76", "d6", "76", "d6", "d6", "76", "d6", "76" };
                    break;
                case 16:
                    columnCodes = new List<string> { "58", "f8", "58", "f8", "58", "f8", "58", "f8", "58", "f8", "58", "58", "f8", "58", "f8" };
                    break;
                case 20:
                    columnCodes = new List<string> { "da", "7a", "da", "7a", "da", "7a", "da", "7a", "da", "7a", "da", "da", "7a", "da", "7a" };
                    break;
                case 24:
                    columnCodes = new List<string> { "dc", "7c", "dc", "7c", "dc", "7c", "dc", "7c", "dc", "7c", "dc", "dc", "7c", "dc", "7c" };
                    break;
                default: // 28
                    columnCodes = new List<string> { "5e", "fe", "5e", "fe", "5e", "fe", "5e", "fe", "5e", "fe", "5e", "5e", "fe", "5e", "fe" };
                    break;
            }
            string code = rowCode + columnCodes[row];

            if (columnRest == 1)
                return code + " " + code + " 97a1 97a1 ";
            if (columnRest == 2)
                return code + " " + code + " 97a2 97a2 ";
            if (columnRest == 3)
                return code + " " + code + " 9723 9723 ";
            return code + " " + code + " ";
        }

        private double GetTimeCodebyText(string code, double startMilliseconds)
        {
            double MICROSECONDS_PER_CODEWORD = 1000.0 * 1000.0 / (Configuration.Settings.General.CurrentFrameRate);

            var lines = code.Trim().SplitToLines();

            double code_words = (code.Length / 2) + 10 + (lines.Count * 2);
            double code_time_microseconds = code_words * MICROSECONDS_PER_CODEWORD;
            double code_start = startMilliseconds * 1000 - code_time_microseconds;
            return code_start / 1000;

        }
        /// <summary>
        /// convert time code based on text for audio and text sync.
        /// </summary>
        /// <param name="code"></param>
        /// <param name="startMilliseconds"></param>
        /// <param name="endMilliseconds"></param>
        /// <param name="previous_end"></param>
        /// <returns></returns>
        private string ToTimeCodebyText(string code, double startMilliseconds, double endMilliseconds, double previous_end = 0.0)
        {
            double code_start = GetTimeCodebyText(code, startMilliseconds);
            return ToTimeCode(code_start);
        }

        private string ToTimeCode(double totalMilliseconds)
        {
            TimeSpan ts = TimeSpan.FromMilliseconds(totalMilliseconds);
            if (DropFrame)
                return string.Format("{0:00}:{1:00}:{2:00};{3:00}", ts.Hours, ts.Minutes, ts.Seconds, MillisecondsToFramesMaxFrameRate(ts.Milliseconds));
            return string.Format("{0:00}:{1:00}:{2:00}:{3:00}", ts.Hours, ts.Minutes, ts.Seconds, MillisecondsToFramesMaxFrameRate(ts.Milliseconds));
        }
        public static SccPositionAndStyle GetColorAndPosition(string code)
        {
            switch (code.ToLower(CultureInfo.InvariantCulture))
            {
                //NO x-coordinate?
                case "1140": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 1, 0);
                case "1160": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 2, 0);
                case "1240": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 3, 0);
                case "1260": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 4, 0);
                case "1540": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 5, 0);
                case "1560": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 6, 0);
                case "1640": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 7, 0);
                case "1660": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 8, 0);
                case "1740": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 9, 0);
                case "1760": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 10, 0);
                case "1040": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 11, 0);
                case "1340": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 12, 0);
                case "1360": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 13, 0);
                case "1440": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 14, 0);
                case "1460": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 15, 0);

                case "1141": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 1, 0);
                case "1161": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 2, 0);
                case "1241": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 3, 0);
                case "1261": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 4, 0);
                case "1541": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 5, 0);
                case "1561": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 6, 0);
                case "1641": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 7, 0);
                case "1661": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 8, 0);
                case "1741": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 9, 0);
                case "1761": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 10, 0);
                case "1041": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 11, 0);
                case "1341": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 12, 0);
                case "1361": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 13, 0);
                case "1441": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 14, 0);
                case "1461": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 15, 0);

                case "1142": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, 1, 0);
                case "1162": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, 2, 0);
                case "1242": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, 3, 0);
                case "1262": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, 4, 0);
                case "1542": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, 5, 0);
                case "1562": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, 6, 0);
                case "1642": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, 7, 0);
                case "1662": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, 8, 0);
                case "1742": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, 9, 0);
                case "1762": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, 10, 0);
                case "1042": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, 11, 0);
                case "1342": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, 12, 0);
                case "1362": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, 13, 0);
                case "1442": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, 14, 0);
                case "1462": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, 15, 0);

                case "1143": return new SccPositionAndStyle(Color.Green, FontStyle.Underline, 1, 0);
                case "1163": return new SccPositionAndStyle(Color.Green, FontStyle.Underline, 2, 0);
                case "1243": return new SccPositionAndStyle(Color.Green, FontStyle.Underline, 3, 0);
                case "1263": return new SccPositionAndStyle(Color.Green, FontStyle.Underline, 4, 0);
                case "1543": return new SccPositionAndStyle(Color.Green, FontStyle.Underline, 5, 0);
                case "1563": return new SccPositionAndStyle(Color.Green, FontStyle.Underline, 6, 0);
                case "1643": return new SccPositionAndStyle(Color.Green, FontStyle.Underline, 7, 0);
                case "1663": return new SccPositionAndStyle(Color.Green, FontStyle.Underline, 8, 0);
                case "1743": return new SccPositionAndStyle(Color.Green, FontStyle.Underline, 9, 0);
                case "1763": return new SccPositionAndStyle(Color.Green, FontStyle.Underline, 10, 0);
                case "1043": return new SccPositionAndStyle(Color.Green, FontStyle.Underline, 11, 0);
                case "1343": return new SccPositionAndStyle(Color.Green, FontStyle.Underline, 12, 0);
                case "1363": return new SccPositionAndStyle(Color.Green, FontStyle.Underline, 13, 0);
                case "1443": return new SccPositionAndStyle(Color.Green, FontStyle.Underline, 14, 0);
                case "1463": return new SccPositionAndStyle(Color.Green, FontStyle.Underline, 15, 0);

                case "1144": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 1, 0);
                case "1164": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 2, 0);
                case "1244": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 3, 0);
                case "1264": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 4, 0);
                case "1544": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 5, 0);
                case "1564": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 6, 0);
                case "1644": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 7, 0);
                case "1664": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 8, 0);
                case "1744": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 9, 0);
                case "1764": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 10, 0);
                case "1044": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 11, 0);
                case "1344": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 12, 0);
                case "1364": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 13, 0);
                case "1444": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 14, 0);
                case "1464": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 15, 0);

                case "1145": return new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 1, 0);
                case "1165": return new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 2, 0);
                case "1245": return new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 3, 0);
                case "1265": return new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 4, 0);
                case "1545": return new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 5, 0);
                case "1565": return new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 6, 0);
                case "1645": return new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 7, 0);
                case "1665": return new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 8, 0);
                case "1745": return new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 9, 0);
                case "1765": return new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 10, 0);
                case "1045": return new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 11, 0);
                case "1345": return new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 12, 0);
                case "1365": return new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 13, 0);
                case "1445": return new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 14, 0);
                case "1465": return new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 15, 0);

                case "1146": return new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 1, 0);
                case "1166": return new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 2, 0);
                case "1246": return new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 3, 0);
                case "1266": return new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 4, 0);
                case "1546": return new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 5, 0);
                case "1566": return new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 6, 0);
                case "1646": return new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 7, 0);
                case "1666": return new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 8, 0);
                case "1746": return new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 9, 0);
                case "1766": return new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 10, 0);
                case "1046": return new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 11, 0);
                case "1346": return new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 12, 0);
                case "1366": return new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 13, 0);
                case "1446": return new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 14, 0);
                case "1466": return new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 15, 0);

                case "1147": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 1, 0);
                case "1167": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 2, 0);
                case "1247": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 3, 0);
                case "1267": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 4, 0);
                case "1547": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 5, 0);
                case "1567": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 6, 0);
                case "1647": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 7, 0);
                case "1667": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 8, 0);
                case "1747": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 9, 0);
                case "1767": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 10, 0);
                case "1047": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 11, 0);
                case "1347": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 12, 0);
                case "1367": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 13, 0);
                case "1447": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 14, 0);
                case "1467": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 15, 0);

                case "1148": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, 1, 0);
                case "1168": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, 2, 0);
                case "1248": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, 3, 0);
                case "1268": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, 4, 0);
                case "1548": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, 5, 0);
                case "1568": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, 6, 0);
                case "1648": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, 7, 0);
                case "1668": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, 8, 0);
                case "1748": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, 9, 0);
                case "1768": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, 10, 0);
                case "1048": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, 11, 0);
                case "1348": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, 12, 0);
                case "1368": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, 13, 0);
                case "1448": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, 14, 0);
                case "1468": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, 15, 0);

                case "1149": return new SccPositionAndStyle(Color.Red, FontStyle.Underline, 1, 0);
                case "1169": return new SccPositionAndStyle(Color.Red, FontStyle.Underline, 2, 0);
                case "1249": return new SccPositionAndStyle(Color.Red, FontStyle.Underline, 3, 0);
                case "1269": return new SccPositionAndStyle(Color.Red, FontStyle.Underline, 4, 0);
                case "1549": return new SccPositionAndStyle(Color.Red, FontStyle.Underline, 5, 0);
                case "1569": return new SccPositionAndStyle(Color.Red, FontStyle.Underline, 6, 0);
                case "1649": return new SccPositionAndStyle(Color.Red, FontStyle.Underline, 7, 0);
                case "1669": return new SccPositionAndStyle(Color.Red, FontStyle.Underline, 8, 0);
                case "1749": return new SccPositionAndStyle(Color.Red, FontStyle.Underline, 9, 0);
                case "1769": return new SccPositionAndStyle(Color.Red, FontStyle.Underline, 10, 0);
                case "1049": return new SccPositionAndStyle(Color.Red, FontStyle.Underline, 11, 0);
                case "1349": return new SccPositionAndStyle(Color.Red, FontStyle.Underline, 12, 0);
                case "1369": return new SccPositionAndStyle(Color.Red, FontStyle.Underline, 13, 0);
                case "1449": return new SccPositionAndStyle(Color.Red, FontStyle.Underline, 14, 0);
                case "1469": return new SccPositionAndStyle(Color.Red, FontStyle.Underline, 15, 0);

                case "114a": return new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 1, 0);
                case "116a": return new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 2, 0);
                case "124a": return new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 3, 0);
                case "126a": return new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 4, 0);
                case "154a": return new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 5, 0);
                case "156a": return new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 6, 0);
                case "164a": return new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 7, 0);
                case "166a": return new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 8, 0);
                case "174a": return new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 9, 0);
                case "176a": return new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 10, 0);
                case "104a": return new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 11, 0);
                case "134a": return new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 12, 0);
                case "136a": return new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 13, 0);
                case "144a": return new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 14, 0);
                case "146a": return new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 15, 0);

                case "114b": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 1, 0);
                case "116b": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 2, 0);
                case "124b": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 3, 0);
                case "126b": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 4, 0);
                case "154b": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 5, 0);
                case "156b": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 6, 0);
                case "164b": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 7, 0);
                case "166b": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 8, 0);
                case "174b": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 9, 0);
                case "176b": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 10, 0);
                case "104b": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 11, 0);
                case "134b": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 12, 0);
                case "136b": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 13, 0);
                case "144b": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 14, 0);
                case "146b": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 15, 0);

                case "114c": return new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 1, 0);
                case "116c": return new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 2, 0);
                case "124c": return new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 3, 0);
                case "126c": return new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 4, 0);
                case "154c": return new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 5, 0);
                case "156c": return new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 6, 0);
                case "164c": return new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 7, 0);
                case "166c": return new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 8, 0);
                case "174c": return new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 9, 0);
                case "176c": return new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 10, 0);
                case "104c": return new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 11, 0);
                case "134c": return new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 12, 0);
                case "136c": return new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 13, 0);
                case "144c": return new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 14, 0);
                case "146c": return new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 15, 0);

                case "114d": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 1, 0);
                case "116d": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 2, 0);
                case "124d": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 3, 0);
                case "126d": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 4, 0);
                case "154d": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 5, 0);
                case "156d": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 6, 0);
                case "164d": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 7, 0);
                case "166d": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 8, 0);
                case "174d": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 9, 0);
                case "176d": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 10, 0);
                case "104d": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 11, 0);
                case "134d": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 12, 0);
                case "136d": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 13, 0);
                case "144d": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 14, 0);
                case "146d": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 15, 0);

                case "114e": return new SccPositionAndStyle(Color.White, FontStyle.Italic, 1, 0);
                case "116e": return new SccPositionAndStyle(Color.White, FontStyle.Italic, 2, 0);
                case "124e": return new SccPositionAndStyle(Color.White, FontStyle.Italic, 3, 0);
                case "126e": return new SccPositionAndStyle(Color.White, FontStyle.Italic, 4, 0);
                case "154e": return new SccPositionAndStyle(Color.White, FontStyle.Italic, 5, 0);
                case "156e": return new SccPositionAndStyle(Color.White, FontStyle.Italic, 6, 0);
                case "164e": return new SccPositionAndStyle(Color.White, FontStyle.Italic, 7, 0);
                case "166e": return new SccPositionAndStyle(Color.White, FontStyle.Italic, 8, 0);
                case "174e": return new SccPositionAndStyle(Color.White, FontStyle.Italic, 9, 0);
                case "176e": return new SccPositionAndStyle(Color.White, FontStyle.Italic, 10, 0);
                case "104e": return new SccPositionAndStyle(Color.White, FontStyle.Italic, 11, 0);
                case "134e": return new SccPositionAndStyle(Color.White, FontStyle.Italic, 12, 0);
                case "136e": return new SccPositionAndStyle(Color.White, FontStyle.Italic, 13, 0);
                case "144e": return new SccPositionAndStyle(Color.White, FontStyle.Italic, 14, 0);
                case "146e": return new SccPositionAndStyle(Color.White, FontStyle.Italic, 15, 0);

                case "114f": return new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 1, 0);
                case "116f": return new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 2, 0);
                case "124f": return new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 3, 0);
                case "126f": return new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 4, 0);
                case "154f": return new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 5, 0);
                case "156f": return new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 6, 0);
                case "164f": return new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 7, 0);
                case "166f": return new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 8, 0);
                case "174f": return new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 9, 0);
                case "176f": return new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 10, 0);
                case "104f": return new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 11, 0);
                case "134f": return new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 12, 0);
                case "136f": return new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 13, 0);
                case "144f": return new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 14, 0);
                case "146f": return new SccPositionAndStyle(Color.White, FontStyle.Underline | FontStyle.Italic, 15, 0);

                case "91d0": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 1, 0);
                case "9151": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 1, 0);
                case "91c2": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, 1, 0);
                case "9143": return new SccPositionAndStyle(Color.Green, FontStyle.Underline, 1, 0);
                case "91c4": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 1, 0);
                case "9145": return new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 1, 0);
                case "9146": return new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 1, 0);
                case "91c7": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 1, 0);
                case "91c8": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, 1, 0);
                case "9149": return new SccPositionAndStyle(Color.Red, FontStyle.Underline, 1, 0);
                case "914a": return new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 1, 0);
                case "91cb": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 1, 0);
                case "914c": return new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 1, 0);
                case "91cd": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 1, 0);

                case "9170": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 2, 0);
                case "91f1": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 2, 0);
                case "9162": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, 2, 0);
                case "91e3": return new SccPositionAndStyle(Color.Green, FontStyle.Underline, 2, 0);
                case "9164": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 2, 0);
                case "91e5": return new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 2, 0);
                case "91e6": return new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 2, 0);
                case "9167": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 2, 0);
                case "9168": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, 2, 0);
                case "91e9": return new SccPositionAndStyle(Color.Red, FontStyle.Underline, 2, 0);
                case "91ea": return new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 2, 0);
                case "916b": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 2, 0);
                case "91ec": return new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 2, 0);
                case "916d": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 2, 0);

                case "92d0": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 3, 0);
                case "9251": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 3, 0);
                case "92c2": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, 3, 0);
                case "9243": return new SccPositionAndStyle(Color.Green, FontStyle.Underline, 3, 0);
                case "92c4": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 3, 0);
                case "9245": return new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 3, 0);
                case "9246": return new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 3, 0);
                case "92c7": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 3, 0);
                case "92c8": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, 3, 0);
                case "9249": return new SccPositionAndStyle(Color.Red, FontStyle.Underline, 3, 0);
                case "924a": return new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 3, 0);
                case "92cb": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 3, 0);
                case "924c": return new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 3, 0);
                case "92cd": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 3, 0);

                case "9270": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 4, 0);
                case "92f1": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 4, 0);
                case "9262": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, 4, 0);
                case "92e3": return new SccPositionAndStyle(Color.Green, FontStyle.Underline, 4, 0);
                case "9264": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 4, 0);
                case "92e5": return new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 4, 0);
                case "92e6": return new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 4, 0);
                case "9267": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 4, 0);
                case "9268": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, 4, 0);
                case "92e9": return new SccPositionAndStyle(Color.Red, FontStyle.Underline, 4, 0);
                case "92ea": return new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 4, 0);
                case "926b": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 4, 0);
                case "92ec": return new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 4, 0);
                case "926d": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 4, 0);

                case "15d0": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 5, 0);
                case "1551": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 5, 0);
                case "15c2": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, 5, 0);
                //                case "1543": return new SCCPositionAndStyle(Color.Green, FontStyle.Underline, 5, 0);
                case "15c4": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 5, 0);
                //                case "1545": return new SCCPositionAndStyle(Color.Blue, FontStyle.Underline, 5, 0);
                //                case "1546": return new SCCPositionAndStyle(Color.Cyan, FontStyle.Regular, 5, 0);
                case "15c7": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 5, 0);
                case "15c8": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, 5, 0);
                //case "1549": return new SCCPositionAndStyle(Color.Red, FontStyle.Underline, 5, 0);
                //case "154a": return new SCCPositionAndStyle(Color.Yellow, FontStyle.Regular, 5, 0);
                case "15cb": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 5, 0);
                //case "154c": return new SCCPositionAndStyle(Color.Magenta, FontStyle.Regular, 5, 0);
                case "15cd": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 5, 0);

                case "1570": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 6, 0);
                case "15f1": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 6, 0);
                //case "1562": return new SCCPositionAndStyle(Color.Green, FontStyle.Regular, 6, 0);
                //case "15e3": return new SCCPositionAndStyle(Color.Green, FontStyle.Underline, 6, 0);
                //case "1564": return new SCCPositionAndStyle(Color.Blue, FontStyle.Regular, 6, 0);
                case "15e5": return new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 6, 0);
                case "15e6": return new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 6, 0);
                //case "1567": return new SCCPositionAndStyle(Color.Cyan, FontStyle.Underline, 6, 0);
                //case "1568": return new SCCPositionAndStyle(Color.Red, FontStyle.Regular, 6, 0);
                case "15e9": return new SccPositionAndStyle(Color.Red, FontStyle.Underline, 6, 0);
                case "15ea": return new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 6, 0);
                //case "156b": return new SCCPositionAndStyle(Color.Yellow, FontStyle.Underline, 6, 0);
                case "15ec": return new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 6, 0);
                //case "156d": return new SCCPositionAndStyle(Color.Magenta, FontStyle.Underline, 6, 0);

                case "16d0": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 7, 0);
                case "1651": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 7, 0);
                case "16c2": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, 7, 0);
                //case "1643": return new SCCPositionAndStyle(Color.Green, FontStyle.Underline, 7, 0);
                case "16c4": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 7, 0);
                //case "1645": return new SCCPositionAndStyle(Color.Blue, FontStyle.Underline, 7, 0);
                //case "1646": return new SCCPositionAndStyle(Color.Cyan, FontStyle.Regular, 7, 0);
                case "16c7": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 7, 0);
                case "16c8": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, 7, 0);
                //case "1649": return new SCCPositionAndStyle(Color.Red, FontStyle.Underline, 7, 0);
                //case "164a": return new SCCPositionAndStyle(Color.Yellow, FontStyle.Regular, 7, 0);
                case "16cb": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 7, 0);
                //case "164c": return new SCCPositionAndStyle(Color.Magenta, FontStyle.Regular, 7, 0);
                case "16cd": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 7, 0);

                case "1670": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 8, 0);
                case "16f1": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 8, 0);
                //case "1662": return new SCCPositionAndStyle(Color.Green, FontStyle.Regular, 8, 0);
                case "16e3": return new SccPositionAndStyle(Color.Green, FontStyle.Underline, 8, 0);
                //case "1664": return new SCCPositionAndStyle(Color.Blue, FontStyle.Regular, 8, 0);
                case "16e5": return new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 8, 0);
                case "16e6": return new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 8, 0);
                //case "1667": return new SCCPositionAndStyle(Color.Cyan, FontStyle.Underline, 8, 0);
                //case "1668": return new SCCPositionAndStyle(Color.Red, FontStyle.Regular, 8, 0);
                case "16e9": return new SccPositionAndStyle(Color.Red, FontStyle.Underline, 8, 0);
                case "16ea": return new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 8, 0);
                //case "166b": return new SCCPositionAndStyle(Color.Yellow, FontStyle.Underline, 8, 0);
                case "16ec": return new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 8, 0);
                //case "166d": return new SCCPositionAndStyle(Color.Magenta, FontStyle.Underline, 8, 0);

                case "97d0": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 9, 0);
                case "9751": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 9, 0);
                case "97c2": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, 9, 0);
                case "9743": return new SccPositionAndStyle(Color.Green, FontStyle.Underline, 9, 0);
                case "97c4": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 9, 0);
                case "9745": return new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 9, 0);
                case "9746": return new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 9, 0);
                case "97c7": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 9, 0);
                case "97c8": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, 9, 0);
                case "9749": return new SccPositionAndStyle(Color.Red, FontStyle.Underline, 9, 0);
                case "974a": return new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 9, 0);
                case "97cb": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 9, 0);
                case "974c": return new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 9, 0);
                case "97cd": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 9, 0);

                case "9770": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 10, 0);
                case "97f1": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 10, 0);
                case "9762": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, 10, 0);
                case "97e3": return new SccPositionAndStyle(Color.Green, FontStyle.Underline, 10, 0);
                case "9764": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 10, 0);
                case "97e5": return new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 10, 0);
                case "97e6": return new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 10, 0);
                case "9767": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 10, 0);
                case "9768": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, 10, 0);
                case "97e9": return new SccPositionAndStyle(Color.Red, FontStyle.Underline, 10, 0);
                case "97ea": return new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 10, 0);
                case "976b": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 10, 0);
                case "97ec": return new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 10, 0);
                case "976d": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 10, 0);

                case "10d0": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 11, 0);
                case "1051": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 11, 0);
                case "10c2": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, 11, 0);
                //case "1043": return new SCCPositionAndStyle(Color.Green, FontStyle.Underline, 11, 0);
                case "10c4": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 11, 0);
                //case "1045": return new SCCPositionAndStyle(Color.Blue, FontStyle.Underline, 11, 0);
                //case "1046": return new SCCPositionAndStyle(Color.Cyan, FontStyle.Regular, 11, 0);
                case "10c7": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 11, 0);
                case "10c8": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, 11, 0);
                //case "1049": return new SCCPositionAndStyle(Color.Red, FontStyle.Underline, 11, 0);
                //case "104a": return new SCCPositionAndStyle(Color.Yellow, FontStyle.Regular, 11, 0);
                case "10cb": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 11, 0);
                //case "104c": return new SCCPositionAndStyle(Color.Magenta, FontStyle.Regular, 11, 0);
                case "10cd": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 11, 0);

                case "13d0": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 12, 0);
                case "1351": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 12, 0);
                case "13c2": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, 12, 0);
                //case "1343": return new SCCPositionAndStyle(Color.Green, FontStyle.Underline, 12, 0);
                case "13c4": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 12, 0);
                //case "1345": return new SCCPositionAndStyle(Color.Blue, FontStyle.Underline, 12, 0);
                //case "1346": return new SCCPositionAndStyle(Color.Cyan, FontStyle.Regular, 12, 0);
                case "13c7": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 12, 0);
                case "13c8": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, 12, 0);
                //case "1349": return new SCCPositionAndStyle(Color.Red, FontStyle.Underline, 12, 0);
                //case "134a": return new SCCPositionAndStyle(Color.Yellow, FontStyle.Regular, 12, 0);
                case "13cb": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 12, 0);
                //case "134c": return new SCCPositionAndStyle(Color.Magenta, FontStyle.Regular, 12, 0);
                case "13cd": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 12, 0);

                case "1370": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 13, 0);
                case "13f1": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 13, 0);
                //case "1362": return new SCCPositionAndStyle(Color.Green, FontStyle.Regular, 13, 0);
                case "13e3": return new SccPositionAndStyle(Color.Green, FontStyle.Underline, 13, 0);
                //case "1364": return new SCCPositionAndStyle(Color.Blue, FontStyle.Regular, 13, 0);
                case "13e5": return new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 13, 0);
                case "13e6": return new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 13, 0);
                //case "1367": return new SCCPositionAndStyle(Color.Cyan, FontStyle.Underline, 13, 0);
                //case "1368": return new SCCPositionAndStyle(Color.Red, FontStyle.Regular, 13, 0);
                case "13e9": return new SccPositionAndStyle(Color.Red, FontStyle.Underline, 13, 0);
                case "13ea": return new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 13, 0);
                //case "136b": return new SCCPositionAndStyle(Color.Yellow, FontStyle.Underline, 13, 0);
                case "13ec": return new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 13, 0);
                //case "136d": return new SCCPositionAndStyle(Color.Magenta, FontStyle.Underline, 13, 0);

                case "94d0": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 14, 0);
                case "9451": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 14, 0);
                case "94c2": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, 14, 0);
                case "9443": return new SccPositionAndStyle(Color.Green, FontStyle.Underline, 14, 0);
                case "94c4": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 14, 0);
                case "9445": return new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 14, 0);
                case "9446": return new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 14, 0);
                case "94c7": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 14, 0);
                case "94c8": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, 14, 0);
                case "9449": return new SccPositionAndStyle(Color.Red, FontStyle.Underline, 14, 0);
                case "944a": return new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 14, 0);
                case "94cb": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 14, 0);
                case "944c": return new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 14, 0);
                case "94cd": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 14, 0);

                case "9470": return new SccPositionAndStyle(Color.White, FontStyle.Regular, 15, 0);
                case "94f1": return new SccPositionAndStyle(Color.White, FontStyle.Underline, 15, 0);
                case "9462": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, 15, 0);
                case "94e3": return new SccPositionAndStyle(Color.Green, FontStyle.Underline, 15, 0);
                case "9464": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, 15, 0);
                case "94e5": return new SccPositionAndStyle(Color.Blue, FontStyle.Underline, 15, 0);
                case "94e6": return new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, 15, 0);
                case "9467": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, 15, 0);
                case "9468": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, 15, 0);
                case "94e9": return new SccPositionAndStyle(Color.Red, FontStyle.Underline, 15, 0);
                case "94ea": return new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, 15, 0);
                case "946b": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, 15, 0);
                case "94ec": return new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, 15, 0);
                case "946d": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, 15, 0);

                //Columns 4-28

                case "9152": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 1, 4);
                case "91d3": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 1, 4);
                case "9154": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 1, 8);
                case "91d5": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 1, 8);
                case "91d6": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 1, 12);
                case "9157": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 1, 12);
                case "9158": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 1, 16);
                case "91d9": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 1, 16);
                case "91da": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 1, 20);
                case "915b": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 1, 20);
                case "91dc": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 1, 24);
                case "915d": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 1, 24);
                case "915e": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 1, 28);
                case "91df": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 1, 28);

                case "91f2": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 2, 4);
                case "9173": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 2, 4);
                case "91f4": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 2, 8);
                case "9175": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 2, 8);
                case "9176": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 2, 12);
                case "91f7": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 2, 12);
                case "91f8": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 2, 16);
                case "9179": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 2, 16);
                case "917a": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 2, 20);
                case "91fb": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 2, 20);
                case "917c": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 2, 24);
                case "91fd": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 2, 24);
                case "91fe": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 2, 28);
                case "917f": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 2, 28);

                case "9252": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 3, 4);
                case "92d3": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 3, 4);
                case "9254": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 3, 8);
                case "92d5": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 3, 8);
                case "92d6": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 3, 12);
                case "9257": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 3, 12);
                case "9258": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 3, 16);
                case "92d9": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 3, 16);
                case "92da": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 3, 20);
                case "925b": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 3, 20);
                case "92dc": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 3, 24);
                case "925d": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 3, 24);
                case "925e": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 3, 28);
                case "92df": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 3, 28);

                case "92f2": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 4, 4);
                case "9273": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 4, 4);
                case "92f4": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 4, 8);
                case "9275": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 4, 8);
                case "9276": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 4, 12);
                case "92f7": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 4, 12);
                case "92f8": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 4, 16);
                case "9279": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 4, 16);
                case "927a": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 4, 20);
                case "92fb": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 4, 20);
                case "927c": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 4, 24);
                case "92fd": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 4, 24);
                case "92fe": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 4, 28);
                case "927f": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 4, 28);

                case "1552": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 5, 4);
                case "15d3": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 5, 4);
                case "1554": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 5, 8);
                case "15d5": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 5, 8);
                case "15d6": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 5, 12);
                case "1557": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 5, 12);
                case "1558": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 5, 16);
                case "15d9": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 5, 16);
                case "15da": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 5, 20);
                case "155b": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 5, 20);
                case "15dc": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 5, 24);
                case "155d": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 5, 24);
                case "155e": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 5, 28);
                case "15df": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 5, 28);

                case "15f2": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 6, 4);
                case "1573": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 6, 4);
                case "15f4": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 6, 8);
                case "1575": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 6, 8);
                case "1576": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 6, 12);
                case "15f7": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 6, 12);
                case "15f8": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 6, 16);
                case "1579": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 6, 16);
                case "157a": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 6, 20);
                case "15fb": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 6, 20);
                case "157c": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 6, 24);
                case "15fd": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 6, 24);
                case "15fe": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 6, 28);
                case "157f": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 6, 28);

                case "1652": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 7, 4);
                case "16d3": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 7, 4);
                case "1654": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 7, 8);
                case "16d5": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 7, 8);
                case "16d6": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 7, 12);
                case "1657": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 7, 12);
                case "1658": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 7, 16);
                case "16d9": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 7, 16);
                case "16da": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 7, 20);
                case "165b": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 7, 20);
                case "16dc": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 7, 24);
                case "165d": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 7, 24);
                case "165e": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 7, 28);
                case "16df": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 7, 28);

                case "16f2": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 8, 4);
                case "1673": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 8, 4);
                case "16f4": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 8, 8);
                case "1675": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 8, 8);
                case "1676": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 8, 12);
                case "16f7": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 8, 12);
                case "16f8": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 8, 16);
                case "1679": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 8, 16);
                case "167a": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 8, 20);
                case "16fb": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 8, 20);
                case "167c": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 8, 24);
                case "16fd": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 8, 24);
                case "16fe": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 8, 28);
                case "167f": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 8, 28);

                case "9752": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 9, 4);
                case "97d3": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 9, 4);
                case "9754": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 9, 8);
                case "97d5": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 9, 8);
                case "97d6": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 9, 12);
                case "9757": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 9, 12);
                case "9758": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 9, 16);
                case "97d9": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 9, 16);
                case "97da": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 9, 20);
                case "975b": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 9, 20);
                case "97dc": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 9, 24);
                case "975d": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 9, 24);
                case "975e": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 9, 28);
                case "97df": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 9, 28);

                case "97f2": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 10, 4);
                case "9773": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 10, 4);
                case "97f4": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 10, 8);
                case "9775": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 10, 8);
                case "9776": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 10, 12);
                case "97f7": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 10, 12);
                case "97f8": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 10, 16);
                case "9779": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 10, 16);
                case "977a": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 10, 20);
                case "97fb": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 10, 20);
                case "977c": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 10, 24);
                case "97fd": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 10, 24);
                case "97fe": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 10, 28);
                case "977f": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 10, 28);

                case "1052": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 11, 4);
                case "10d3": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 11, 4);
                case "1054": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 11, 8);
                case "10d5": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 11, 8);
                case "10d6": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 11, 12);
                case "1057": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 11, 12);
                case "1058": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 11, 16);
                case "10d9": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 11, 16);
                case "10da": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 11, 20);
                case "105b": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 11, 20);
                case "10dc": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 11, 24);
                case "105d": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 11, 24);
                case "105e": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 11, 28);
                case "10df": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 11, 28);

                case "1352": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 12, 4);
                case "13d3": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 12, 4);
                case "1354": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 12, 8);
                case "13d5": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 12, 8);
                case "13d6": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 12, 12);
                case "1357": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 12, 12);
                case "1358": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 12, 16);
                case "13d9": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 12, 16);
                case "13da": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 12, 20);
                case "135b": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 12, 20);
                case "13dc": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 12, 24);
                case "135d": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 12, 24);
                case "135e": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 12, 28);
                case "13df": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 12, 28);

                case "13f2": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 13, 4);
                case "1373": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 13, 4);
                case "13f4": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 13, 8);
                case "1375": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 13, 8);
                case "1376": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 13, 12);
                case "13f7": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 13, 12);
                case "13f8": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 13, 16);
                case "1379": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 13, 16);
                case "137a": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 13, 20);
                case "13fb": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 13, 20);
                case "137c": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 13, 24);
                case "13fd": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 13, 24);
                case "13fe": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 13, 28);
                case "137f": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 13, 28);

                case "9452": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 14, 4);
                case "94d3": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 14, 4);
                case "9454": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 14, 8);
                case "94d5": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 14, 8);
                case "94d6": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 14, 12);
                case "9457": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 14, 12);
                case "9458": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 14, 16);
                case "94d9": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 14, 16);
                case "94da": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 14, 20);
                case "945b": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 14, 20);
                case "94dc": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 14, 24);
                case "945d": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 14, 24);
                case "945e": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 14, 28);
                case "94df": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 14, 28);

                case "94f2": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 15, 4);
                case "9473": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 15, 4);
                case "94f4": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 15, 8);
                case "9475": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 15, 8);
                case "9476": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 15, 12);
                case "94f7": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 15, 12);
                case "94f8": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 15, 16);
                case "9479": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 15, 16);
                case "947a": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 15, 20);
                case "94fb": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 15, 20);
                case "947c": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 15, 24);
                case "94fd": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 15, 24);
                case "94fe": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, 15, 28);
                case "947f": return new SccPositionAndStyle(Color.Transparent, FontStyle.Underline, 15, 28);

                // mid-row commands
                case "9120": return new SccPositionAndStyle(Color.White, FontStyle.Regular, -1, -1);
                case "91a1": return new SccPositionAndStyle(Color.White, FontStyle.Underline, -1, -1);
                case "91a2": return new SccPositionAndStyle(Color.Green, FontStyle.Regular, -1, -1);
                case "9123": return new SccPositionAndStyle(Color.Green, FontStyle.Underline, -1, -1);
                case "91a4": return new SccPositionAndStyle(Color.Blue, FontStyle.Regular, -1, -1);
                case "9125": return new SccPositionAndStyle(Color.Blue, FontStyle.Underline, -1, -1);
                case "9126": return new SccPositionAndStyle(Color.Cyan, FontStyle.Regular, -1, -1);
                case "91a7": return new SccPositionAndStyle(Color.Cyan, FontStyle.Underline, -1, -1);
                case "91a8": return new SccPositionAndStyle(Color.Red, FontStyle.Regular, -1, -1);
                case "9129": return new SccPositionAndStyle(Color.Red, FontStyle.Underline, -1, -1);
                case "912a": return new SccPositionAndStyle(Color.Yellow, FontStyle.Regular, -1, -1);
                case "91ab": return new SccPositionAndStyle(Color.Yellow, FontStyle.Underline, -1, -1);
                case "912c": return new SccPositionAndStyle(Color.Magenta, FontStyle.Regular, -1, -1);
                case "91ad": return new SccPositionAndStyle(Color.Magenta, FontStyle.Underline, -1, -1);

                case "91ae": return new SccPositionAndStyle(Color.Transparent, FontStyle.Italic, -1, -1);
                case "912f": return new SccPositionAndStyle(Color.Transparent, FontStyle.Italic | FontStyle.Underline, -1, -1);

                case "94a8": return new SccPositionAndStyle(Color.Transparent, FontStyle.Regular, -1, -1); // turn flash on
            }
            return null;
        }
        // get the code from postion, color fontstyle
        private static string GetCodeFromPositionFormats(Color color, FontStyle fontStyle, int y, int x)
        {
            SccPositionAndStyle sccPosition = new SccPositionAndStyle(color, fontStyle, y, x);
            //var letter = PositionDictionary.FirstOrDefault(n => n.Value.X == sccPosition.X && n.Value.Y == sccPosition.Y && n.Value.ForeColor == sccPosition.ForeColor && n.Value.Style == sccPosition.Style);
            var letter = PositionDictionary.FirstOrDefault(n => n.Value.X == sccPosition.X && n.Value.Y == sccPosition.Y);
            if (letter.Equals(new KeyValuePair<string, SccPositionAndStyle>()))
                return null;
            return letter.Key;
        }
        public override void LoadSubtitle(Subtitle subtitle, List<string> lines, string fileName)
        {
            _errorCount = 0;
            Paragraph p = null;
            foreach (string line in lines)
            {
                string s = line.Trim();
                var match = RegexTimeCodes.Match(s);
                if (match.Success)
                {
                    TimeCode startTime = ParseTimeCode(s.Substring(0, match.Length - 1));
                    string text = GetSccText(s.Substring(match.Index), ref _errorCount);

                    if (text == "942c 942c" || text == "942c")
                    {
                        if (p != null)
                        {
                            p.EndTime = new TimeCode(startTime.TotalMilliseconds);
                        }
                    }
                    else
                    {
                        p = new Paragraph(startTime, new TimeCode(startTime.TotalMilliseconds), text);
                        subtitle.Paragraphs.Add(p);
                    }
                }
            }
            for (int i = subtitle.Paragraphs.Count - 2; i >= 0; i--)
            {
                p = subtitle.GetParagraphOrDefault(i);
                Paragraph next = subtitle.GetParagraphOrDefault(i + 1);
                if (p != null && next != null && Math.Abs(p.EndTime.TotalMilliseconds - p.StartTime.TotalMilliseconds) < 0.001)
                    p.EndTime = new TimeCode(next.StartTime.TotalMilliseconds);
                if (next != null && string.IsNullOrEmpty(next.Text))
                    subtitle.Paragraphs.Remove(next);
            }
            p = subtitle.GetParagraphOrDefault(0);
            if (p != null && string.IsNullOrEmpty(p.Text))
                subtitle.Paragraphs.Remove(p);

            subtitle.Renumber();
        }

        public static string GetSccText(string s, ref int errorCount)
        {
            int y = 0;
            string[] parts = s.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            var sb = new StringBuilder();
            bool first = true;
            bool italicOn = false;
            int k = 0;
            while (k < parts.Length)
            {
                string part = parts[k];
                if (part.Length == 4)
                {
                    if (part != "94ae" && part != "9420" && part != "94ad" && part != "9426" && part != "946e" && part != "91ce" && part != "13ce" && part != "9425" && part != "9429")
                    {

                        // skewed apos "’"
                        if (part == "9229" && k < parts.Length - 1 && parts[k + 1] == "9229" && sb.EndsWith('\''))
                        {
                            sb.Remove(sb.Length - 1, 1);
                            sb.Append("’");
                            k += 2;
                            continue;
                        }

                        // 3 codes
                        if (k < parts.Length - 2)
                        {
                            var letter = GetLetterFromCode(part + " " + parts[k + 1] + " " + parts[k + 2]);
                            if (letter != null)
                            {
                                sb.Append(letter);
                                k += 3;
                                continue;
                            }
                        }

                        // two codes
                        if (k < parts.Length - 1)
                        {
                            var letter = GetLetterFromCode(part + " " + parts[k + 1]);
                            if (letter != null)
                            {
                                sb.Append(letter);
                                k += 2;
                                continue;
                            }
                        }

                        if (part[0] == '9' || part[0] == '8')
                        {
                            if (k + 1 < parts.Length && parts[k + 1] == part)
                                k++;
                        }

                        var cp = GetColorAndPosition(part);
                        if (cp != null)
                        {
                            if (cp.Y > 0 && y >= 0 && cp.Y > y && !sb.ToString().EndsWith(Environment.NewLine, StringComparison.Ordinal) && !string.IsNullOrWhiteSpace(sb.ToString()))
                                sb.AppendLine();
                            if (cp.Y > 0)
                                y = cp.Y;
                            if ((cp.Style & FontStyle.Italic) == FontStyle.Italic && !italicOn)
                            {
                                sb.Append("<i>");
                                italicOn = true;
                            }
                            else if (cp.Style == FontStyle.Regular && italicOn)
                            {
                                sb.Append("</i>");
                                italicOn = false;
                            }
                        }
                        else
                        {
                            switch (part)
                            {
                                case "9440":
                                case "94e0":
                                    if (!sb.ToString().EndsWith(Environment.NewLine, StringComparison.Ordinal))
                                        sb.AppendLine();
                                    break;
                                case "2c75":
                                case "2cf2":
                                case "2c6f":
                                case "2c6e":
                                case "2c6d":
                                case "2c6c":
                                case "2c6b":
                                case "2c6a":
                                case "2c69":
                                case "2c68":
                                case "2c67":
                                case "2c66":
                                case "2c65":
                                case "2c64":
                                case "2c63":
                                case "2c62":
                                case "2c61":
                                    sb.Append(GetLetterFromCode(part.Substring(2, 2)));
                                    break;
                                case "2c52":
                                case "2c94":
                                    break;
                                default:
                                    var result = GetLetterFromCode(part);
                                    if (result == null)
                                    {
                                        sb.Append(GetLetterFromCode(part.Substring(0, 2)));
                                        var secondPart = part.Substring(2, 2) + "80";
                                        var foundSecondPart = false;

                                        // 3 codes
                                        if (k < parts.Length - 2)
                                        {
                                            var letter = GetLetterFromCode(secondPart + " " + parts[k + 1] + " " + parts[k + 2]);
                                            if (letter != null)
                                            {
                                                sb.Append(letter);
                                                k += 3;
                                                continue;
                                            }
                                        }

                                        // two codes
                                        if (k < parts.Length - 1 && !foundSecondPart)
                                        {
                                            var letter = GetLetterFromCode(secondPart + " " + parts[k + 1]);
                                            if (letter != null)
                                            {
                                                sb.Append(letter);
                                                k += 2;
                                                continue;
                                            }
                                        }

                                        sb.Append(GetLetterFromCode(part.Substring(2, 2)));
                                    }
                                    else
                                    {
                                        sb.Append(result);
                                    }
                                    break;
                            }
                        }
                    }
                }
                else if (part.Length > 0)
                {
                    if (!first)
                        errorCount++;
                }
                first = false;
                k++;
            }
            string res = sb.ToString().Replace("<i></i>", string.Empty).Replace("</i><i>", string.Empty);
            res = res.Replace("  ", " ").Replace("  ", " ").Replace(Environment.NewLine + " ", Environment.NewLine).Trim();
            if (res.Contains("<i>") && !res.Contains("</i>"))
                res += "</i>";
            return HtmlUtil.FixInvalidItalicTags(res);
        }

        private static TimeCode ParseTimeCode(string start)
        {
            string[] arr = start.Split(new[] { ':', ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
            return new TimeCode(int.Parse(arr[0]), int.Parse(arr[1]), int.Parse(arr[2]), FramesToMillisecondsMax999(int.Parse(arr[3])));
        }

    }
}