using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DokuWikiFormatter
{
    public sealed class DWikiSyntax
    {
        public enum ListOrder
        {
            Ordered = 0,
            Unordered = 1,
        }

        private static string sntx_BoldText_b = "**";
        private static string sntx_BoldText_e = "**";
        private static string sntx_ItalicText_b = "//";
        private static string sntx_ItalicText_e = "//";
        private static string sntx_UnderLinedText_b = "__";
        private static string sntx_UnderLinedText_e = "__";
        private static string sntx_StrikeThrough_b = "<del>";
        private static string sntx_StrikeThrough_e = "</del>";
        private static string sntc_Line = "----";

        private static string sntx_Heading1 = "======";
        private static string sntx_Heading2 = "=====";
        private static string sntx_Heading3 = "====";
        private static string sntx_Heading4 = "===";
        private static string sntx_Heading5 = "==";

        private static string sntx_UnorderedList = "*";
        private static string sntx_OrderedList = "-";

        private static string sntx_TwoWhitespace = "  ";
        private static string sntx_Whitespace = " ";

        private static string sntx_InternalLink_b = "[[";
        private static string sntx_InternalLink_e = "]]";
        private static string sntx_InternalLink_s = "|";

        private static string sntx_ExternalLink_b = "[[";
        private static string sntx_ExternalLink_e = "]]";
        private static string sntx_ExternalLink_s = "|";

        private static string sntx_TableHeader = "^";
        private static string sntx_TableCell = "|";
        private static string sntx_TableNewLine = @" \\ ";
        private static string sntx_TableCell_VMerge = ":::";

        private static string sntx_IEImage_b = "{{";
        private static string sntx_IEImage_e = "}}";
        private static string sntx_ImageSizeAdd = "?";
        private static string sntx_ImageSizeSeparator = "x";

        private static string sntx_NoFormatting_b = "<nowiki>";
        private static string sntx_NoFormatting_e = "</nowiki>";

        public static string BoldText_b { get => sntx_BoldText_b; }
        public static string BoldText_e { get => sntx_BoldText_e; }
        public static string ItalicText_b { get => sntx_ItalicText_b; }
        public static string ItalicText_e { get => sntx_ItalicText_e; }
        public static string UnderLinedText_b { get => sntx_UnderLinedText_b; }
        public static string UnderLinedText_e { get => sntx_UnderLinedText_e; }
        public static string StrikeThrough_b { get => sntx_StrikeThrough_b; }
        public static string StrikeThrough_e { get => sntx_StrikeThrough_e; }
        public static string Heading1 { get => sntx_Heading1;}
        public static string Heading2 { get => sntx_Heading2; }
        public static string Heading3 { get => sntx_Heading3; }
        public static string Heading4 { get => sntx_Heading4; }
        public static string Heading5 { get => sntx_Heading5;  }
        public static string UnorderedList { get => sntx_UnorderedList; }
        public static string OrderedList { get => sntx_OrderedList;}
        public static string TwoWhitespace { get => sntx_TwoWhitespace; }
        public static string InternalLink_b { get => sntx_InternalLink_b; }
        public static string InternalLink_e { get => sntx_InternalLink_e; }
        public static string InternalLink_s { get => sntx_InternalLink_s; }
        public static string ExternalLink_b { get => sntx_ExternalLink_b; }
        public static string ExternalLink_e { get => sntx_ExternalLink_e; }
        public static string ExternalLink_s { get => sntx_ExternalLink_s;}
        public static string TableHeader { get => sntx_TableHeader; }
        public static string TableCell { get => sntx_TableCell; }
        public static string TableNewLine { get => sntx_TableNewLine; }
        public static string TableCell_VMerge { get => sntx_TableCell_VMerge; }
        public static string IEImage_b { get => sntx_IEImage_b; }
        public static string IEImage_e { get => sntx_IEImage_e; }
        public static string NoFormatting_b { get => sntx_NoFormatting_b; }
        public static string NoFormatting_e { get => sntx_NoFormatting_e; }
        public static string Whitespace { get => sntx_Whitespace; }
        public static string Line { get => sntc_Line; }
        public static string ImageSizeAdd { get => sntx_ImageSizeAdd; }
        public static string ImageSizeSeparator { get => sntx_ImageSizeSeparator; }
    }
}
