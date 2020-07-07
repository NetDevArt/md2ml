using System.ComponentModel;

namespace Md2Ml.Enum
{
    /// <summary>
    /// Enumerator of styles to be defined in your docx template
    /// In order to apply those styles when converting markdown elements to docx elements via openXML
    /// </summary>
    public enum DocStyles
    {
        [Description("Heading1")]
        Heading1,
        [Description("Heading2")]
        Heading2,
        [Description("Heading3")]
        Heading3,
        [Description("Heading4")]
        Heading4,
        [Description("Heading5")]
        Heading5,
        [Description("Heading6")]
        Heading6,
        [Description("Heading7")]
        Heading7,
        [Description("Heading8")]
        Heading8,
        [Description("Heading9")]
        Heading9,
        [Description("CodeBlock")]
        CodeBlock,
        [Description("Quote1")]
        Quote
    }

    public static class DocStylesExtensions
    {
        // Get the [Description("...")] text by its value
        public static string ToDescriptionString(this DocStyles val)
        {
            DescriptionAttribute[] attributes = (DescriptionAttribute[])val
                .GetType()
                .GetField(val.ToString())
                .GetCustomAttributes(typeof(DescriptionAttribute), false);
            return attributes.Length > 0 ? attributes[0].Description : string.Empty;
        }
    }
}