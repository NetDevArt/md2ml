using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Md2Ml.Enum;
using A = DocumentFormat.OpenXml.Drawing;
using GridColumn = DocumentFormat.OpenXml.Wordprocessing.GridColumn;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using Path = System.IO.Path;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using TableCellProperties = DocumentFormat.OpenXml.Wordprocessing.TableCellProperties;
using TableGrid = DocumentFormat.OpenXml.Wordprocessing.TableGrid;
using TableProperties = DocumentFormat.OpenXml.Wordprocessing.TableProperties;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using TableStyle = DocumentFormat.OpenXml.Wordprocessing.TableStyle;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using Underline = DocumentFormat.OpenXml.Wordprocessing.Underline;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace Md2Ml
{
	public class Md2MlEngine : IDisposable
	{
		private WordprocessingDocument _package;
		private MainDocumentPart _mainDocumentPart;
		private Document _document;
		private Body _body;
        private string _fileDirectory;

        /// <summary>
        /// Create or overwrite a document docx (openXml sdk)
        /// Build the main structure of the document 
        /// </summary>
        /// <param name="templateName">Path to the template document used</param>
        /// <param name="fileName">Path to the rendered document</param>
		public void CreateDocument(string templateName, string fileName)
        {
            if (File.Exists(fileName))
                File.Delete(fileName);
            File.Copy(templateName, fileName);

			_package = WordprocessingDocument.Open(fileName, true);
			_mainDocumentPart = _package.MainDocumentPart;
            _document = _mainDocumentPart.Document;
            _body = _document.Body;
            _body.RemoveAllChildren<Paragraph>();
			
		}

		/// <summary>
		/// This method should be used if your image paths are relative to the directory of the parsed file
		/// </summary>
		/// <param name="filepath">The path of the markdown file</param>
        public void SetFileDirectory(string filepath)
        {
            _fileDirectory = Path.GetDirectoryName(filepath);
        }
        public string GetFileDirectory()
        {
            return _fileDirectory;
        }

		#region Table Supports
		public Table CreateTable(int cols)
        {
            Table table = new Table();
            TableProperties tableProperties = new TableProperties(
                new TableStyle() {Val = "Professionnel"},
				new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct });
            table.AppendChild<TableProperties>(tableProperties);
			
            // May be useless, or for performances reasons
            TableGrid tg = new TableGrid();
            for (int i = 0; i < cols; i++)
                tg.Append(new GridColumn());
			table.AppendChild(tg);
			_body.Append(table);
            return table;
            
		}

		/// <summary>
		/// Add a row to a table, defining of not the justification for each cell
		/// </summary>
		/// <param name="table">The table on which append a TableRow</param>
		/// <param name="values">List of values to be inserted col by col in a row</param>
		/// <param name="justif">The list of justification value</param>
        public void AddTableRow(Table table, List<string> values, List<JustificationValues> justif = null)
        {
            TableRow tableRow = new TableRow();
            int i = 0;
            foreach (var value in values)
            {
                TableCell tableCell = new TableCell();
                TableCellProperties tableCellProperties = new TableCellProperties();
                TableCellWidth tableCellWidth = new TableCellWidth() { Type = TableWidthUnitValues.Auto };
                tableCellProperties.Append(tableCellWidth);
                tableCell.Append(tableCellProperties);
                tableRow.Append(tableCell);

                var para = justif != null ? CreateNonBodyParagraph(justif[i]) : CreateNonBodyParagraph();
                
                WriteText(para, value.Trim());
                tableCell.Append(para);
                
                i++;
            }
            table.Append(tableRow);
        }
		#endregion

		#region Lists support
		/// <summary>
		/// TODO : Problem to fix, all lists are represented as Numbered...
		/// Process an array of markdown items of a list (with spaces to define the item level)
		/// The pattern can be detected as a CodeBlock pattern, so fix it with the try catch.
		/// I should probably not let defining a CodeBlock with spaces and tabulations,
		/// but only with tabulations... (WIP)
        /// </summary>
		/// <param name="core"></param>
		/// <param name="bulletedItems">Items of a list, split in an array by break lines</param>
		/// <param name="paragraphStyle"></param>
		public void MarkdownList(Md2MlEngine core, string[] bulletedItems, string paragraphStyle = "ParagraphList")
        {
            foreach (var item in bulletedItems)
            {
				// Detect if item is ordered or not
                var matchedPattern = PatternMatcher.GetMarkdownMatch(item);
                if (matchedPattern.Key != ParaPattern.OrderedList && matchedPattern.Key != ParaPattern.UnorderedList)
                {
                    try
                    {
                        matchedPattern = PatternMatcher.GetMatchFromPattern(item, ParaPattern.OrderedList);
                    }
                    catch (Exception e)
                    {
						matchedPattern = PatternMatcher.GetMatchFromPattern(item, ParaPattern.UnorderedList);
					}
                }
                // Then count spaces 3 by 3 to define the level of the item
                var nbSpaces = matchedPattern.Value.Groups[0].Value.TakeWhile(Char.IsWhiteSpace).Count(); ;
                var itemLvl = nbSpaces / 3;
				// Then Create paragraph, properties and format the text
                Paragraph paragraph1 = CreateParagraph(paragraphStyle);
                NumberingProperties numberingProperties1 = new NumberingProperties();
                NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() { Val = itemLvl };
                NumberingId numberingId1 = GetListType(matchedPattern.Key);
				
				numberingProperties1.Append(numberingLevelReference1);
                numberingProperties1.Append(numberingId1);
                paragraph1.ParagraphProperties.Append(numberingProperties1);
                MarkdownStringParser.FormatText(core, paragraph1, matchedPattern.Value.Groups[2].Value, new StyleProperties());
            }
        }

        private NumberingId GetListType(ParaPattern pattern)
        {
			if(pattern == ParaPattern.OrderedList)
				return new NumberingId() { Val = 2 };
            else
                return new NumberingId() { Val = 1 };
		}
		#endregion

		#region Paragraph supports
		/// <summary>
		/// Create a void XML paragraph object
		/// </summary>
		/// <returns>The paragraph where to insert some stuff</returns>
		public Paragraph CreateParagraph()
        {
            var Para = new Paragraph();
            _body.Append(Para);
            return Para;
        }
        public Paragraph CreateParagraph(ParaProperties properties)
        {
            ParagraphProperties paraProp = new ParagraphProperties();
            if (properties.StyleName != null) paraProp.Append(new ParagraphStyleId() { Val = properties.StyleName });
            if (properties.Alignment != JustificationValues.Left) paraProp.Append(new Justification() { Val = properties.Alignment });
            Indentation ind = new Indentation();
            if (properties.FirstLineIndent == 0) ind.FirstLine = (properties.FirstLineIndent * 567).ToString("n0").Replace(",", "");
            if (properties.LeftIndent == 0) ind.FirstLine = (properties.LeftIndent * 567).ToString("n0").Replace(",", "");
            if (properties.RightIndent == 0) ind.FirstLine = (properties.RightIndent * 567).ToString("n0").Replace(",", "");
            var para = CreateParagraph(); para.Append(paraProp); return para;
        }
        public Paragraph CreateParagraph(string paragraphStyleName)
        {
            if (string.IsNullOrEmpty(paragraphStyleName))
                return CreateParagraph();
            
            ParagraphProperties paraProp = new ParagraphProperties();
            paraProp.Append(new ParagraphStyleId() { Val = paragraphStyleName });
            var para = new Paragraph();
            para.Append(paraProp);
            _body.Append(para);
            return para;
        }
		private Paragraph CreateNonBodyParagraph() => new Paragraph();
		public Paragraph CreateNonBodyParagraph(JustificationValues alignment)
        {
            if (alignment == JustificationValues.Left) return CreateNonBodyParagraph();

            ParagraphProperties paraProp = new ParagraphProperties();
            paraProp.Append(new Justification() { Val = alignment });
            var para = CreateNonBodyParagraph();
            para.Append(paraProp);
            return para;
        }

		/// <summary>
		/// Write a text in a paragraph, with some styles
		/// </summary>
		/// <param name="paragraph"></param>
		/// <param name="text"></param>
		/// <param name="fontProperties"></param>
		public void WriteText(Paragraph paragraph, string text, StyleProperties fontProperties)
        {
            Run run = new Run();
            RunProperties rp = new RunProperties();
            if (fontProperties.StyleName != null)
                rp.Append(new RunStyle() { Val = fontProperties.StyleName });
            if (fontProperties.FontName != null)
                rp.Append(new RunFonts() { ComplexScript = fontProperties.FontName, Ascii = fontProperties.FontName, HighAnsi = fontProperties.FontName });
            else if (fontProperties.UseTemplateHeadingFont)
                rp.Append(new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, ComplexScriptTheme = ThemeFontValues.MajorHighAnsi });
            if (fontProperties.FontSize != null) { rp.Append(new FontSize() { Val = fontProperties.FontSize }); }
            if (fontProperties.Bold) rp.Append(new Bold());
            if (fontProperties.Italic) rp.Append(new Italic());
            if (fontProperties.Underline != UnderlineValues.None) rp.Append(new Underline() { Val = fontProperties.Underline });
            if (fontProperties.Strikeout) rp.Append(new Strike());
            if (fontProperties.WriteAs != VerticalPositionValues.Baseline) rp.Append(new VerticalTextAlignment() { Val = fontProperties.WriteAs });
            if (fontProperties.UseThemeColor) rp.Append(new DocumentFormat.OpenXml.Wordprocessing.Color() { ThemeColor = fontProperties.ThemeColor });
            else if (fontProperties.Color != null) rp.Append(new DocumentFormat.OpenXml.Wordprocessing.Color() { Val = string.Format("#{0:X2}{1:X2}{2:X2}", fontProperties.Color.Value.R, fontProperties.Color.Value.G, fontProperties.Color.Value.B) });
            run.Append(rp);
            run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            paragraph.Append(run);
        }
        public void WriteText(string text) => WriteText(CreateParagraph(), text);
        public void WriteText(Paragraph paragraph, string text) => WriteText(paragraph, text, new StyleProperties());

		/// <summary>
		/// Transform the markdown content to openXML (docx) format.
		/// </summary>
		/// <param name="content">The markdown content to transform</param>
		public void WriteMdText(string content) => MarkdownStringParser.Parse(this, content);
        #endregion

		#region Support for IMAGES insertion (with original dimensions, and in paragraph or not)
		/// <summary>
		/// Add an image to the body by its stream
		/// </summary>
		/// <param name="image"></param>
		public void AddImage(Stream image)
		{
			ImagePart imagePart = _mainDocumentPart.AddImagePart("image/png");
			imagePart.FeedData(image);
			AddImageToBody(_package, _mainDocumentPart.GetIdOfPart(imagePart));
		}

		/// <summary>
		/// Add an image (by its path on the system) to the body of the document by getting its dimensions
        /// </summary>
		/// <param name="absoluteImgPath">The absolute path of the image</param>
		public void AddImage(string absoluteImgPath, Paragraph para = null)
		{
			var dimensions = GetDimensions(absoluteImgPath);
			var imagePart = GetImagePart(absoluteImgPath);
			AddImageToBody(_package, _mainDocumentPart.GetIdOfPart(imagePart), para, dimensions);
		}
		private ImagePart GetImagePart(string path)
		{
			using (FileStream stream = new FileStream(path, FileMode.Open))
			{
				ImagePart imagePart = _mainDocumentPart.AddImagePart("image/png");
				imagePart.FeedData(stream);
				return imagePart;
			}
		}

		/// <summary>
		/// Get dimensions of an image by its absolute path on the system
		/// </summary>
		/// <param name="path">The path to the image to read the file as a stream</param>
		/// <returns>A tuple representing image dimensions as pixels and dpi</returns>
		private ((int width, int height) pixel, (int width, int height) dpi) GetDimensions(string path)
		{
			using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read))
			{
				using (Image tif = Image.FromStream(stream: fs,
				useEmbeddedColorManagement: false,
				validateImageData: false))
				{
					float pixelWidth = tif.PhysicalDimension.Width;
					float pixelHeight = tif.PhysicalDimension.Height;
					var pixelDim = (width: Convert.ToInt32(pixelWidth), height: Convert.ToInt32(pixelHeight));

					float dpiX = tif.HorizontalResolution;
					float dpiY = tif.VerticalResolution;
					var dpiDim = (width: Convert.ToInt32(dpiX), height: Convert.ToInt32(dpiY));

					return (pixel: pixelDim, dpi: dpiDim);
				}
			}
		}

		/// <summary>
		/// <summary>
		/// Add an image to a paragraph
		/// </summary>
		/// <param name="document">The WordprocessingDocument created where to insert the image</param>
		/// <param name="relationshipId"></param>
		/// <param name="para"></param>
		/// <param name="dimImg">Optional: If set, the image will be inserted with its own dimension</param
		/// </summary>
		private static void AddImageToBody(WordprocessingDocument document, string relationshipId, Paragraph para = null, ((int width, int height) pixel, (int width, int height) dpi) dimImg = default(((int width, int height), (int width, int height))))
		{
			// Default image size
			var widthEmus = 990000L;
			var heightEmus = 792000L;

			// If dimensions are passed in params, compute the image size within document size before inserting the image
			if (dimImg != default)
			{
				var widthPx = dimImg.pixel.width;
				var heightPx = dimImg.pixel.height;
				var hRezDpi = dimImg.dpi.width;
				var vRezDpi = dimImg.dpi.height;
				const int emusPerInch = 914400;
				const int emusPerCm = 360000;
				var maxWidthCm = 16; // Width per cm of your word document
				widthEmus = (long)(widthPx / hRezDpi * emusPerInch);
				heightEmus = (long)(heightPx / vRezDpi * emusPerInch);
				var maxWidthEmus = (long)(maxWidthCm * emusPerCm);
				if (widthEmus > maxWidthEmus)
				{
					var ratio = (heightEmus * 1.0m) / widthEmus;
					widthEmus = maxWidthEmus;
					heightEmus = (long)(widthEmus * ratio);
				}
				// Perhaps add condition if images are too small
			}

			var element =
				 new Drawing(
					 new Wp.Inline(
						 new Wp.Extent() { Cx = widthEmus, Cy = heightEmus },
						 new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
						 new Wp.DocProperties() { Id = (UInt32Value)1U, Name = "Picture 1" },
						 new Wp.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
						 new A.Graphic(
							 new A.GraphicData(
								 new Pic.Picture(
									 new Pic.NonVisualPictureProperties(
										 new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "New Bitmap Image.Png" },
										 new Pic.NonVisualPictureDrawingProperties()),
									 new Pic.BlipFill(
										 new A.Blip(
											 new A.BlipExtensionList(
												 new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" })
										 )
										 {
											 Embed = relationshipId,
											 CompressionState = A.BlipCompressionValues.Print
										 },
										 new A.Stretch(
											 new A.FillRectangle())),
									 new Pic.ShapeProperties(
										 new A.Transform2D(
											 new A.Offset() { X = 0L, Y = 0L },
											 new A.Extents() { Cx = widthEmus, Cy = heightEmus }),
										 new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }))
							 )
							 { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
					 )
					 {
						 DistanceFromTop = (UInt32Value)0U,
						 DistanceFromBottom = (UInt32Value)0U,
						 DistanceFromLeft = (UInt32Value)0U,
						 DistanceFromRight = (UInt32Value)0U,
						 EditId = "50D07946"
					 });

            // Append the reference to body, the element should be in a Run.
			Run childElement = new Run(element);
            if (para == null)
            {
                para = new Paragraph(childElement);
                document.MainDocumentPart.Document.Body.AppendChild(para);
			}
			else 
				para.Append(childElement);

		}
        #endregion


		public void Cleanup(OpenXmlElement element) => element.RemoveAllChildren();
        public void SaveDocument() => _package.Save();
		public void SaveDocument(string fileName) => _package.SaveAs(fileName);

		#region IDisposable Support
		private bool disposedValue = false; // To detect redundant calls
		protected virtual void Dispose(bool disposing)
		{
			SaveDocument();
			if (!disposedValue)
			{
				if (disposing)
				{
					_package.Dispose();
				}
				_package = null;
				disposedValue = true;
			}
		}
		public void Dispose() => Dispose(true);
		#endregion
	}
}
