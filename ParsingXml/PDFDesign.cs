using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Security;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsumingFaysalBankService
{
    public class PDFDesign
    {
        //The MigraDoc Section,document, paragraph and table
        Section section;
        Document document;
        Paragraph paragraph;
        Table table;

        #region Some Pre-Defined colors
        readonly public static Color TableBorder = new Color(81, 125, 192);
        readonly public static Color TableBlue = new Color(235, 240, 249);
        readonly public static Color TableGray = new Color(203, 203, 200);
        readonly public static Color DarkBlue = new Color(0, 2, 45);
        readonly public static Color Blue = new Color(39, 35, 141);
        readonly public static Color Green = new Color(0, 112, 120);
        #endregion

        public bool IsTableExist { get; set; }

        //Creates the document.
        public bool CreateDocumentForTransHist(
            bool isTableExist,
            List<string> paragraphList,
            TableData tableHeaderData,
            List<List<string>> tableBodyData,
            TableData topTablesHeaderData,
            List<List<TopTable>> topTables,
            string filePath,
            string imagePath,
            string OpeningBalanceDate,
            string ClosingBalanceDate,
            string password = null)
        {
            try
            {
                IsTableExist = isTableExist;
                string pdfTitle = paragraphList[0];
                string openingBalance = paragraphList[1];
                string closingBalance = paragraphList[2];
                paragraphList.RemoveRange(0, 3);

                this.document = new Document();
                section = this.document.AddSection();
                this.section.PageSetup.LeftMargin = 20;
                this.section.PageSetup.RightMargin = 20;
                this.section.PageSetup.TopMargin = 30;
                DefineStyles();

                //CreateImage($"{imagePath}\\FaysalBankLogo.png", "130");
                CreateImage($"{imagePath}\\header.jpg", "19.6cm");

                LineBreak(1);

                #region Top Heading
                Paragraph header = section.AddParagraph(pdfTitle);
                header.Style = StyleNames.Header;
                #endregion

                LineBreak(1);

                this.table = section.AddTable();
                this.table.Style = "Table";
                this.table.Borders.Color = Colors.White;

                CreateTableHeader(topTablesHeaderData);

                //TopTable(topTables);

                CreateImage($"{imagePath}\\New Islamic Logo.PNG", "19.6cm");

                paragraph = section.AddParagraph();
                paragraph.Style = "P1";
                paragraph.AddText("Faysal Bank Limited has revised its schedule of charges from 01 Jan 2023. For details please contact your branch or call (021) 111 06 06 06");

                Paragraphs(paragraphList);

                // Second right aligned paragraph
                //paragraph = section.AddParagraph();
                //paragraph.Style = "RightAlignedTitle";
                //paragraph.AddText($"{openingBalance}");

                // Second left aligned paragraph
                //paragraph = section.AddParagraph();
                //paragraph.Format.Alignment = ParagraphAlignment.Left;
                //paragraph.AddSpace(50);
                //paragraph.AddText($"Opening Balance as on: {OpeningBalanceDate}");
                //paragraph.Format.LeftIndent = 0.5;

                if (IsTableExist)
                {
                    this.table = section.AddTable();
                    this.table.Style = "Table";
                    this.table.RightPadding = 2;
                    this.table.LeftPadding = 2;
                    this.table.TopPadding = 4;
                    this.table.BottomPadding = 4;
                    this.table.Borders.Color = Colors.Gray;
                    this.table.Borders.Width = 0.25;
                    this.table.Borders.Left.Width = 0.5;
                    this.table.Borders.Right.Width = 0.5;
                    this.table.Rows.LeftIndent = 0;
                    CreateTableHeader(tableHeaderData);
                }


                FillContent(tableBodyData, closingBalance, ClosingBalanceDate);
                PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true);
                pdfRenderer.Document = document;
                pdfRenderer.RenderDocument();

                #region Make PDF Password Protected
                if (!String.IsNullOrEmpty(password))
                {
                    MemoryStream ms = new MemoryStream();
                    pdfRenderer.Save(ms, false);
                    byte[] buffer = new byte[ms.Length];
                    ms.Seek(0, SeekOrigin.Begin);
                    ms.Flush();
                    ms.Read(buffer, 0, (int)ms.Length);
                    pdfRenderer = new PdfDocumentRenderer(true);
                    PdfDocument secureDoc = PdfReader.Open(ms);
                    PdfSecuritySettings securitySettings = secureDoc.SecuritySettings;

                    // Setting one of the passwords automatically sets the security level to 
                    // PdfDocumentSecurityLevel.Encrypted128Bit.
                    securitySettings.UserPassword = password;
                    securitySettings.OwnerPassword = password;

                    // Don't use 40 bit encryption unless needed for compatibility
                    //securitySettings.DocumentSecurityLevel = PdfDocumentSecurityLevel.Encrypted40Bit;

                    // Restrict some rights.
                    securitySettings.PermitAccessibilityExtractContent = false;
                    securitySettings.PermitAnnotations = false;
                    securitySettings.PermitAssembleDocument = false;
                    securitySettings.PermitExtractContent = false;
                    securitySettings.PermitFormsFill = true;
                    securitySettings.PermitFullQualityPrint = false;
                    securitySettings.PermitModifyDocument = true;
                    securitySettings.PermitPrint = false;
                    ms.Close();
                    pdfRenderer.PdfDocument = secureDoc;
                }
                #endregion

                pdfRenderer.Save(filePath);
                return true;
            }
            catch (Exception ex)
            {
                //IOLogger.LogFile.LogGeneralException(ex.ToString(), LogType.Error);
                //ex.ToString();
                Console.WriteLine(ex.Message);
                return false;
            }
        }
        

        /// <summary>
        /// Defines the styles used to format the MigraDoc document.
        /// </summary>
        void DefineStyles()
        {
            // Get the predefined style Normal.
            Style style = this.document.Styles["Normal"];
            style.Font.Size = 10;
            style.Font.Name = "Calibri";
            // Because all styles are derived from Normal, the next line changes the 
            // font of the whole document. Or, more exactly, it changes the font of
            // all styles and paragraphs that do not redefine the font.

            //Styles for document header
            style = this.document.Styles[StyleNames.Header];
            style.ParagraphFormat.AddTabStop("16cm", TabAlignment.Right);
            style.Font.Size = "14";
            style.Font.Color = Green;
            style.Font.Bold = true;

            //Style for document footer
            style = this.document.Styles[StyleNames.Footer];
            style.ParagraphFormat.Shading.Color = Green;
            style.ParagraphFormat.Font.Color = Colors.White;
            style.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            // Create a new style called Table based on style Normal
            style = this.document.Styles.AddStyle("TableCell", "Normal");
            style.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            style.ParagraphFormat.LineSpacing = 20;
            style.ParagraphFormat.LineSpacingRule = LineSpacingRule.AtLeast;

            // Create a new style called RightAlignedTitle based on style Normal
            style = this.document.Styles.AddStyle("RightAlignedTitle", StyleNames.Normal);
            style.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            style.ParagraphFormat.SpaceAfter = new Unit(-11, UnitType.Point);

            // Create a new style called Paragraph based on style Normal
            style = this.document.Styles.AddStyle("Paragraph", "Normal");
            style.ParagraphFormat.Font.Color = Colors.Black;
            style.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            style.ParagraphFormat.LineSpacingRule = LineSpacingRule.AtLeast;
            style.ParagraphFormat.Font.Bold = true;

            // Create a new style called P1 based on style Paragraph
            style = this.document.Styles.AddStyle("P1", "Paragraph");
            style.ParagraphFormat.Shading.Color = Green;
            style.ParagraphFormat.Font.Color = Colors.White;
            style.ParagraphFormat.Font.Size = "11";
            style.ParagraphFormat.SpaceAfter = 10;
            style.ParagraphFormat.SpaceBefore = 10;
            style.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        }

        public Image CreateImage(string path, string width)
        {
            Image image = this.section.AddImage(path);
            image.LockAspectRatio = true;
            image.Width = width;
            return image;
        }

        public void LineBreak(int lineBreaks)
        {
            if (this.paragraph == null)
            {
                this.paragraph = section.AddParagraph();
            }
            for (int i = 0; i < lineBreaks; i++)
            {
                this.paragraph.AddLineBreak();
            }
        }

        public void TopTable(List<List<TopTable>> list)
        {
            foreach (var lstItem in list)
            {
                Row row1 = this.table.AddRow();
                for (int i = 0; i < list[0].Count; i++)
                {
                    row1.Cells[i].AddParagraph(lstItem[i].Text);
                    row1.Cells[i].Style = "TableCell";
                    row1.Cells[i].Format.Font.Bold = lstItem[i].isBold;
                    row1.Cells[i].Format.Font.Color = lstItem[i].Color;
                }
                this.table.SetEdge(0, 0, row1.Cells.Count, 1, Edge.Box, BorderStyle.Single, 0.75);
            }
        }

        public void Paragraphs(List<string> list)
        {
            int i = 0;
            foreach (var item in list)
            {
                int spaceAfter = 5;
                if (i == 1)
                {
                    spaceAfter = 15;
                }
                else if (i == 3)
                {
                    spaceAfter = 25;
                }
                paragraph = section.AddParagraph();
                paragraph.Style = (i == 0) ? "Paragraph" : "Normal";
                paragraph.Format.SpaceBefore = 5;
                paragraph.Format.SpaceAfter = spaceAfter;
                paragraph.AddText(item);
                i++;
            }
        }

        //public void ParagraphsTransactionHistory(List<List<string>> list)
        //{
        //    int i, j = 0;
        //    try
        //    {
        //        for (i = 0; i < list.Count; i++)
        //        {
        //            for (j = 0; j < 3; j++)
        //            {
        //                int spaceAfter = 5;
        //                if (i == 1)
        //                {
        //                    spaceAfter = 5;
        //                }
        //                else if (i == 3)
        //                {
        //                    spaceAfter = 10;
        //                }
        //                paragraph = section.AddParagraph();
        //                paragraph.Style = (i == 0) ? "Paragraph" : "Normal";
        //                paragraph.Format.SpaceBefore = 5;
        //                paragraph.Format.SpaceAfter = spaceAfter;
        //                paragraph.Format.Font.Bold = true;
        //                paragraph.AddText(list[i][j]);
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {

        //        throw;
        //    }



        //}

        public void FixedParagraph(List<string> list)
        {
            TextFormat textFormat = TextFormat.Bold;
            foreach (var item in list)
            {
                LineBreak(3);
                paragraph.AddFormattedText(item, textFormat);
                paragraph.Format.Font.Color = Green;
                paragraph.Format.Font.Size = "11";
                textFormat = TextFormat.NoUnderline;
            }
        }

        public void CreateTableHeader(TableData tableData)
        {
            Column column;
            foreach (var item in tableData.TableProps)
            {
                column = this.table.AddColumn(item.ColumnSize);
                column.Format.Alignment = item.ColumnAlign;
            }
            Row row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = tableData.RowAlignment;
            row.Format.Font.Bold = true;
            row.Shading.Color = tableData.RowColor;
            int count = 0;
            foreach (var item in tableData.TableProps)
            {

                row.Cells[count].AddParagraph(item.ColumnName);
                row.Cells[count].Format.Alignment = item.ColumnAlign;
                row.Cells[count].Format.Font.Color = item.ColumnColor;
                count++;
            }
            this.table.SetEdge(0, 0, count, 1, Edge.Box, BorderStyle.Single, 0.75, Color.Empty);
        }

        public void CreateTableBody(List<List<string>> list)
        {
            foreach (var lstItem in list)
            {
                Row row1 = this.table.AddRow();
                for (int i = 0; i < list[0].Count; i++)
                {
                    row1.Cells[i].AddParagraph(lstItem[i]);
                    row1.Cells[i].VerticalAlignment = VerticalAlignment.Center;
                    row1.Cells[0].Format.Alignment = ParagraphAlignment.Right;
                    row1.Cells[2].Format.Alignment = ParagraphAlignment.Left;
                }
                this.table.SetEdge(0, 0, row1.Cells.Count, 1, Edge.Box, BorderStyle.Single, 0.75);
            }
        }

        /// <summary>
        /// Creates the dynamic parts of the document.
        /// </summary>
        void FillContent(List<List<string>> list, string closingBalance, string closingBalanceDate)
        {
            if (IsTableExist)
            {
                CreateTableBody(list);
            }

            // Second right aligned paragraph
            //paragraph = section.AddParagraph();
            //paragraph.Style = "RightAlignedTitle";
            //paragraph.AddText($"PKR {closingBalance}");

            //// Second left aligned paragraph
            //paragraph = section.AddParagraph();
            //paragraph.Format.Alignment = ParagraphAlignment.Left;
            //paragraph.AddSpace(50);
            //paragraph.AddText($"Closing Balance as on: {closingBalanceDate}");
            //paragraph.Format.LeftIndent = 0.5;

            paragraph = section.AddParagraph();

            LineBreak(5);

            FixedParagraph(new List<string>() {
                "Disclaimer Statement.",
                "This communication/Message (including any attachments) may contain information that is privileged and/or confidential under applicable laws. If you are not the intended recipient or such recipient s employee or agent, you are hereby notified that any dissemination, copy or disclosure of this communication is strictly prohibited. If you have recieved this communication in error, please immediately notify via return internet email to the sender and delete this communication without making any copy.Any unauthorized use or communication of this communication / message in whole or in part is strictly prohibited. Please note that emails are susceptible to change.",
                "Faysal Bank Limited, shall not be liable for the improper or incomplete transmission of the information contained in this communication nor for any delay in its receipt or damage to your system.",
                "Faysal Bank Limited does not guarantee that the integrity of this communication has been maintained nor that this communication is free of viruses, interceptions or interference.",
                "Any unauthorized form of reproduction of this message may result in disciplinary and legal action."
            });
            #region Page Footer
            paragraph = section.Footers.Primary.AddParagraph();
            paragraph.Format.SpaceBefore = -40;
            paragraph.AddText("Page ");
            paragraph.AddPageField();
            paragraph.AddText(" of ");
            paragraph.AddNumPagesField();
            paragraph.Format.Font.Bold = true;
            paragraph = section.Footers.Primary.AddParagraph();
            paragraph.AddText("For further details visit us at ");
            paragraph.AddHyperlink("https://www.faysalbank.com/", HyperlinkType.Web).AddFormattedText("www.faysalbank.com");
            LineBreak(2);
            section.Footers.EvenPage.Add(paragraph.Clone());
            #endregion
        }
    }

    public class TopTable
    {
        public string Text { get; set; }
        public bool isBold { get; set; }
        public Color Color { get; set; }
        public ParagraphAlignment ColumnAlign { get; set; }
    }
    public class TableDefaultProps
    {
        public string ColumnSize { get; set; }
        public ParagraphAlignment ColumnAlign { get; set; }
        public Color ColumnColor { get; set; }
        public Color ColumnBgColor { get; set; }
    }
    public class TableProps : TableDefaultProps
    {
        public TableProps(TableDefaultProps tableDefaultProps)
        {
            this.ColumnSize = tableDefaultProps.ColumnSize;
            this.ColumnAlign = tableDefaultProps.ColumnAlign;
            this.ColumnColor = tableDefaultProps.ColumnColor;
            this.ColumnBgColor = tableDefaultProps.ColumnBgColor;
        }
        public string ColumnName { get; set; }
    }
    public class TableData
    {
        public List<TableProps> TableProps { get; set; }
        public Color RowColor { get; set; }
        public ParagraphAlignment RowAlignment { get; set; }

    }
}
