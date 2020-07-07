# Md2OpenXML
Convert Markdown text in WordProcessingML using OpenXML specifications.  
A work done by forking [md2ml from ashok19r91d](https://github.com/ashok19r91d/md2ml).

# Usage Instructions
	using (Md2OpenXml.Md2OpenXmlCore engine = new Md2OpenXml.Md2OpenXmlCore())
	{
	  engine.CreateDocument(@"D:\<your_path>\template.docx", @"D:\<your_path>\NewDocumentMd2Ml.docx");
	  engine.SetFileDirectory(@"D:\<your_absolute_path>\<processed_file>.<ext>");
	  engine.WriteMdText("# Intro\nGo ahead, play around with the editor! Be sure to check out **bold** and *italic* styling, or even [links](https://google.com). You can type the Markdown syntax, use the toolbar, or use shortcuts like `cmd-b` or `ctrl-b`.\n\n## Lists\nUnordered lists can be started using the toolbar or by typing `*`, `-`, or `+`. Ordered lists can be started by typing `1. `.\n\n### Unordered\n* Lists are a piece of cake\n* They even auto continue as you type\n* A double enter will end them\n* Tabs and shift-tabs work too\n\n### Ordered\n1. Numbered lists...\n2. ...work too!\n\n## What about images?\n![Yes](https://i.imgur.com/sZlktY7.png)\n| Ashok | Arun RD | Himalaya |\n|:-|:-:|-:|\n|As | Ar | 12 |\n|As | Ar | 12 |");
	}

Be careful, the `engine.SetFileDirectory` must be set when the path of an image in markdown is relative to this (markdown) file.

## Create your template.docx
OpenXML specifications let us using some "paragraph" styles to styling each markdown element.  
In order to fit with my developments :

Create a template.docx file:
* It will be used as template when creating document
* Headers, footers, and styles will be copied to your destination file

## Create your own styles
* Open your created template.docx
* Create some new styles
   * The name of your styles should refer to elements in Enum.DocStyles `[Description("Heading1")]`


# Extended Usage Instructions:
Apart from writing Markdown Text to Worddocument this document also let you draft Word Document from the ground.

| Function | Description |
|---|---|
|`CreateParagraph`| Creates a Paragraph attach it to `Document`'s body |
|`CreateNonBodyParagraph`| Create a Paragraph without attaching it to `Document`. Is this useful for creating tables, pagebreaks etc.,|
|`WriteText`| Writes a plain text to document |
|`WriteMdText`| Writes Markdown formatted text to document. Regardless of the type of markdown element, you would rather to use this method  |
|`Cleanup`| Clear entire document |
|`SaveDocument`| Save documnet |
