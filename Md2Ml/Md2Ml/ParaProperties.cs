using DocumentFormat.OpenXml.Wordprocessing;

namespace Md2Ml
{
	public class ParaProperties
	{
		public decimal FirstLineIndent = 0;
		public decimal LeftIndent = 0;
		public decimal RightIndent = 0;
		public JustificationValues Alignment = JustificationValues.Left;
		public string StyleName = null;
	}
}
