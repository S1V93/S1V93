
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocToPDFConverter;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Graphics;
using Syncfusion.Pdf.Parsing;
using Syncfusion.XlsIO;
using Syncfusion.ExcelToPdfConverter;
using System;
using System.Drawing;
using System.IO;
					
public class Program
{
	public static void Main()
	{
		//Get the path of documents into DirectoryInfo
		DirectoryInfo directoryInfo = new DirectoryInfo("../../Documents/");
 
		//Create a new PDF document
        PfDocument document = new PdfDocument();
 
		//Initialize MemoryStream
		MemoryStream stream = new MemoryStream();
 
	//Convert the documents to PDF
foreach (FileInfo fileInfo in directoryInfo.GetFiles())
{
        //Import pages from PDF document and store it into stream
        if (fileInfo.Extension == ".pdf")
            stream = ImportPagesFromPDF(fileInfo.FullName);
 
        //Convert the Excel document into PDF and store it in stream
        if (fileInfo.Extension == ".xls" || fileInfo.Extension == ".xlsx")
            stream = ConvertExcelToPDF(fileInfo.FullName);
 
        //Convert the Word document into PDF and store it in stream
        if (fileInfo.Extension == ".doc" || fileInfo.Extension == ".docx")
            stream = ConvertWordToPDF(fileInfo.FullName);
 
        //Convert the Image file into PDF and store it in stream
        if (fileInfo.Extension == ".jpg" || fileInfo.Extension == ".png" || fileInfo.Extension == ".jpeg")
            stream = ImageToPDF(fileInfo.FullName);
 
        //Load the stream into PdfLoadedDocument
        PdfLoadedDocument loadedDocument = new PdfLoadedDocument(stream);
 
        //Add the watermark
        AddWatermark(loadedDocument, document);
}
 
       //Save the PDF document
          document.Save("MergedPDF.pdf");
 
//Close the instance of PdfDocument
// document.Close(true);
// document.Close(true);
	}
	
	public MemoryStream ConvertWordToPDF(string filePath)
	{
		//Load an existing Word document into WordDocument
WordDocument wordDocument = new WordDocument(filePath, FormatType.Automatic);
 
//Initialize the DocToPDFConverter
DocToPDFConverter converter = new DocToPDFConverter();
 
//Convert Word document into PDF document
PdfDocument document = converter.ConvertToPDF(wordDocument);
 
//Save the PDF document as a stream
MemoryStream stream = new MemoryStream();
document.Save(stream);
 
//Close the instances of PdfDocument and WordDocument
document.Close(true);
wordDocument.Close();
 
//Return the stream
return stream;
	}
	
	public void AddWatermark(PdfLoadedDocument PdfLoadedDocument ,PdfDocument document )
	{
		for (int i = 0; i < c.Pages.Count; i++)
{
    //Set the width, height, and margins for PdfDocument
    document.PageSettings.Margins.All = 0;
    document.PageSettings.Width = 500;
    document.PageSettings.Height = 800;
 
    //Create a PdfTemplate
    PdfTemplate template = loadedDocument.Pages[i].CreateTemplate();
 
    //Add page to the PdfDocument
    PdfPage page = document.Pages.Add();
 
    //Create PdfGraphics for the page
    PdfGraphics graphics = page.Graphics;
 
    //Draw the PDF template
    graphics.DrawPdfTemplate(template, new PointF(0, 0), new SizeF(page.GetClientSize().Width, page.GetClientSize().Height));
 
    //Initialize PdfStandardFont
    PdfFont font = new PdfStandardFont(PdfFontFamily.Helvetica, 14, PdfFontStyle.Bold);
 
    //Save the graphics of the page
    PdfGraphicsState state = page.Graphics.Save();
 
    //Draw the watermark on PDF page
    graphics.SetTransparency(0.5f);
    graphics.DrawString("Page Number : "+ document.Pages.Count.ToString(), font, PdfBrushes.Red, new PointF(350, 400));
}
	}
}