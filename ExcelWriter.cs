using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

using System.Web;

using OpenExcel.OfficeOpenXml.Style;
using OpenExcel.OfficeOpenXml;
namespace grid_excel_net
{
    public class ExcelWriter
    {
        private ExcelDocument wb;
	    private ExcelWorksheet sheet;
	    private ExcelColumn[][] cols;
	    private int colsNumber = 0;
	    private ExcelXmlParser parser;
	
	    public int headerOffset = 0;
	    public int scale = 6;
	    public String pathToImgs = "";//optional, physical path

	    String bgColor = "";
	    String lineColor = "";
	    String headerTextColor = "";
	    String scaleOneColor = "";
	    String scaleTwoColor = "";
	    String gridTextColor = "";
	    String watermarkTextColor = "";

	    private int cols_stat;
	    private int rows_stat;
	  //  RGBColor colors;
	    private String watermark = null;
	
	    public void generate(String xml, HttpResponse resp){
            generate(xml, resp.OutputStream);
            outputExcel(resp);
	    }


        public void generate(string xml, Stream output){
            parser = new ExcelXmlParser();
            //    try {
            parser.setXML(xml);
            createExcel(output);
            setColorProfile();
            headerPrint(parser);
            wb.Workbook.Document.Styles.Save();
            rowsPrint(parser, output);
            footerPrint(parser);
            insertHeader(parser, output);
            insertFooter(parser, output);
            watermarkPrint(parser);
          //  wb.Workbook.Document.Styles.Save();
            wb.Dispose();
            //   } catch (Exception e) {
            //	    e.printStackTrace();
            //   }
        }

	    private void createExcel(Stream resp){
		    /* Save generated excel to file.
		     * Can be useful for debug output.
		     * */
		    /*
		    FileOutputStream fos = new FileOutputStream("d:/test.xls");
		    wb = Workbook.createWorkbook(fos);
		    */
		    wb = ExcelDocument.CreateWorkbook(resp);
            
		    sheet = wb.Workbook.Worksheets.Add("First Sheet");
            wb.EnsureStylesDefined();
		  //  colors = new RGBColor();
	    }

	    private void outputExcel(HttpResponse resp){
		    resp.ContentType = "application/vnd.ms-excel";
		    resp.HeaderEncoding = Encoding.UTF8;
		    resp.AppendHeader("Content-Disposition", "attachment;filename=grid.xls");
		    resp.AppendHeader("Cache-Control", "max-age=0");
		  //  wb.write();
		   // wb.close();
	    }

	    private void headerPrint(ExcelXmlParser parser){
		    cols = parser.getColumnsInfo("head");
		
		    int[] widths = parser.getWidths();
		    this.cols_stat = widths.Length;
		
		    int sumWidth = 0;
		    for (int i = 0; i < widths.Length; i++) {
			    sumWidth += widths[i];
		    }
           
		    if (parser.getWithoutHeader() == false) {
                ExcelFont font = wb.CreateFont("Arial", 9);
                font.Bold = true;
                

			    for (uint row = 1; row <= cols.Length; row++) {
                  
                    sheet.Rows[row].Height = 22.5;
				    for (uint col = 1; col <= cols[row-1].Length; col++) {
                  
                        sheet.Columns[col].Width = widths[col-1] / scale;
					    String name = cols[row-1][col-1].GetName();
                       

                        sheet.Cells[row, col].Style.Font = font;

                        ///TODO: 
                        ///border color, font color, merge cells, alignment
                        sheet.Cells[row, col].Style.Border.BottomStyle = OpenExcel.OfficeOpenXml.Style.ExcelBorderStyleValues.Thin;
                        sheet.Cells[row, col].Style.Border.LeftStyle = OpenExcel.OfficeOpenXml.Style.ExcelBorderStyleValues.Thin;
                        sheet.Cells[row, col].Style.Border.RightStyle = OpenExcel.OfficeOpenXml.Style.ExcelBorderStyleValues.Thin;
                        sheet.Cells[row, col].Style.Border.TopStyle = OpenExcel.OfficeOpenXml.Style.ExcelBorderStyleValues.Thin;
                        sheet.Cells[row, col].Value = name;
                        sheet.Cells[row, col].Style.Fill.ForegroundColor = bgColor;
					    colsNumber = (int)col;
				    }
			    }
			    headerOffset = cols.Length;
			  /*  for (int col = 0; col < cols.Length; col++) {
				    for (int row = 0; row < cols[col].Length; row++) {
					    int cspan = cols[col][row].GetColspan();
					    if (cspan > 0) {
						    sheet.mergeCells(row, col, row + cspan - 1, col);
					    }
					    int rspan = cols[col][row].GetRowspan();
					    if (rspan > 0) {
						    sheet.mergeCells(row, col, row, col + rspan - 1);
					    }
				    }
			    }*/
		    }
	    }

	    private void footerPrint(ExcelXmlParser parser){
		    cols = parser.getColumnsInfo("foot");

		    if (parser.getWithoutHeader() == false) {
                ExcelFont font = wb.CreateFont("Arial", 10);
                font.Bold = true;
			    for (uint row = 1; row <= cols.Length; row++) {
                    sheet.Rows[row].Height = 22.5;
                    uint rowInd = (uint)(row + headerOffset);
                 
				    for (uint col = 1; col <= cols[row-1].Length; col++) {
					
                        sheet.Cells[rowInd, col].Style.Fill.ForegroundColor = bgColor;
                        sheet.Cells[rowInd, col].Style.Font = font;

                        //TODO add text color, vertical alignment, border line color, horizontal alignment

                        sheet.Cells[rowInd, col].Style.Border.BottomStyle = OpenExcel.OfficeOpenXml.Style.ExcelBorderStyleValues.Thin;
                        sheet.Cells[rowInd, col].Style.Border.LeftStyle = OpenExcel.OfficeOpenXml.Style.ExcelBorderStyleValues.Thin;
                        sheet.Cells[rowInd, col].Style.Border.RightStyle = OpenExcel.OfficeOpenXml.Style.ExcelBorderStyleValues.Thin;
                        sheet.Cells[rowInd, col].Style.Border.TopStyle = OpenExcel.OfficeOpenXml.Style.ExcelBorderStyleValues.Thin;

                        sheet.Cells[rowInd, col].Value = cols[row - 1][col - 1].GetName();
				    }
			    }
			  /*  for (int col = 0; col < cols.Length; col++) {
				    for (int row = 0; row < cols[col].Length; row++) {
					    int cspan = cols[col][row].GetColspan();
					    if (cspan > 0) {
						    sheet.mergeCells(row, headerOffset + col, row + cspan - 1, headerOffset + col);
					    }
					    int rspan = cols[col][row].GetRowspan();
					    if (rspan > 0) {
						    sheet.mergeCells(row, headerOffset + col, row, headerOffset + col + rspan - 1);
					    }
				    }
			    }*/
		    }
		    headerOffset += cols.Length;
	    }

	    private void watermarkPrint(ExcelXmlParser parser){
		    if (watermark == null) return;
         //   ExcelFont font = wb.CreateFont("Arial", 10);
         //   font.Bold = true;
         //   font.Color = watermarkTextColor;
        //    foreach(var i in 
           // sheet.Cells[headerOffset + 1, 0
		   /* WritableFont font = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD);
		    font.setColour(colors.getColor(watermarkTextColor, wb));
		    WritableCellFormat f = new WritableCellFormat (font);
		    f.setBorder(Border.ALL, BorderLineStyle.THIN, colors.getColor(lineColor, wb));
		    f.setVerticalAlignment(VerticalAlignment.CENTRE);

		    f.setAlignment(Alignment.CENTRE);
		    Label label = new Label(0, headerOffset, watermark , f);
		    sheet.addCell(label);
		    sheet.mergeCells(0, headerOffset, colsNumber, headerOffset);*/
	    }

	    private void rowsPrint(ExcelXmlParser parser, Stream resp) {
		    //do we really need them?
		    ExcelRow[] rows = parser.getGridContent();
		    this.rows_stat = rows.Length;

          //  ExcelFont font = wb.CreateFont("Arial", 9);
          //  font.Bold = true;

		    for (uint row = 1; row <= rows.Length; row++) {
			    ExcelCell[] cells = rows[row-1].getCells();
                uint rowInd = (uint)(row + headerOffset);
                sheet.Rows[rowInd].Height = 20;
			  //  sheet.Rows[(uint)col].Height = (col + headerOffset + 400);
			    for (uint col = 1; col <= cells.Length; col++) {
				    // sets cell font


                  //  sheet.Cells[rowInd, row].Style.Border.BottomStyle.
                    if (row == rows.Length - 1)
                    {
                        sheet.Cells[rowInd, col].Style.Border.BottomStyle = OpenExcel.OfficeOpenXml.Style.ExcelBorderStyleValues.Thin;                  
                    }
                    if (col == cols.Length - 1)
                    {
                        sheet.Cells[rowInd, col].Style.Border.LeftStyle = OpenExcel.OfficeOpenXml.Style.ExcelBorderStyleValues.Thin;
                    }

                    sheet.Cells[rowInd, col].Style.Border.RightStyle = OpenExcel.OfficeOpenXml.Style.ExcelBorderStyleValues.Thin;
                    sheet.Cells[rowInd, col].Style.Border.TopStyle = OpenExcel.OfficeOpenXml.Style.ExcelBorderStyleValues.Thin;
                    sheet.Cells[rowInd, col].Value = cells[col - 1].GetValue();
              

                    ExcelFont font = wb.CreateFont("Arial", 10);
                    font.Bold = cells[col - 1].GetBold();
                    font.Italic = cells[col - 1].GetItalic();
                   // if ((!cells[col - 1].GetTextColor().Equals("")) && (parser.getProfile().Equals("full_color")))
                  //      font.Color = "FF" + cells[row].GetTextColor();
                  //  else
                   //     font.Color = "FFFF00FF";

                    sheet.Cells[rowInd, col].Style.Font = font;


				  /*  WritableFont font = new WritableFont(WritableFont.ARIAL, 10, (cells[row].getBold()) ? WritableFont.BOLD : WritableFont.NO_BOLD, (cells[row].getItalic()) ? true : false);
				    if ((!cells[row].getTextColor().equals(""))&&(parser.getProfile().equals("full_color")))
					    font.setColour(colors.getColor(cells[row].getTextColor(), wb));
				    else
					    font.setColour(colors.getColor(gridTextColor, wb));
				    WritableCellFormat f = new WritableCellFormat (font);
                    */
				    // sets cell background color
                    
				    if ((!cells[col - 1].GetBgColor().Equals(""))&&(parser.getProfile().Equals("full_color"))) {
                        sheet.Cells[rowInd, col].Style.Fill.ForegroundColor = "FF" + cells[col - 1].GetBgColor();
				    } else {
					    //Colour bg;
                        if (row % 2 == 0)
                        {
                            sheet.Cells[rowInd, col].Style.Fill.ForegroundColor = scaleTwoColor;
						
					    } else {
                            sheet.Cells[rowInd, col].Style.Fill.ForegroundColor = scaleOneColor;
					    }
				    }
                    /*
				    

				    String al = cells[row].getAlign();
				    if (al == "")
					    al = cols[0][row].getAlign();
				    if (al.equalsIgnoreCase("left")) {
					    f.setAlignment(Alignment.LEFT);
				    } else {
					    if (al.equalsIgnoreCase("right")) {
						    f.setAlignment(Alignment.RIGHT);
					    } else {
						    f.setAlignment(Alignment.CENTRE);
					    }
				    }
				    try {
					    double name = Double.parseDouble(cells[row].getValue());
					    Number label = new Number(row, col + headerOffset, name, f);
					    sheet.addCell(label);
				    } catch (Exception e) {
					    String name = cells[row].getValue();
					    Label label = new Label(row, col + headerOffset, name, f);
					    sheet.addCell(label);
				    }*/
			    }
		    }
		    headerOffset += rows.Length;
	    }

	    private void insertHeader(ExcelXmlParser parser, Stream resp){
		   /* if (parser.getHeader() == true) {
			    sheet.insertRow(0);
			    sheet.setRowView(0, 5000);
			    File imgFile = new File(pathToImgs + "/header.png");
			    WritableImage img = new WritableImage(0, 0, cols[0].length, 1, imgFile);
			    sheet.addImage(img);
			    headerOffset++;
		    }*/
           // sheet.
	    }

	    private void insertFooter(ExcelXmlParser parser, Stream resp) {
		 /*   if (parser.getFooter() == true) {
			    sheet.setRowView(headerOffset, 5000);
			    File imgFile = new File(pathToImgs + "/footer.png");
			    WritableImage img = new WritableImage(0, headerOffset, cols[0].length, 1, imgFile);
			    sheet.addImage(img);
		    }*/
	    }

	    public int getColsStat() {
		    return this.cols_stat;
	    }
	
	    public int getRowsStat() {
		    return this.rows_stat;
	    }

	    private void setColorProfile() {
            var alpha = "FF";
		    String profile = parser.getProfile();
		    if ((profile.ToLower().Equals("color"))||profile.ToLower().Equals("full_color")) {
                bgColor = alpha + "D1E5FE";
                lineColor = alpha +  "A4BED4";
                headerTextColor =alpha +  "000000";
                scaleOneColor = alpha + "FFFFFF";
                scaleTwoColor = alpha + "E3EFFF";
                gridTextColor = alpha + "00FF00";
                watermarkTextColor = alpha + "8b8b8b";
		    } else {
			    if (profile.ToLower().Equals("gray")) {
                    bgColor =alpha +  "E3E3E3";
                    lineColor = alpha + "B8B8B8";
                    headerTextColor = alpha + "000000";
                    scaleOneColor = alpha + "FFFFFF";
                    scaleTwoColor = alpha + "EDEDED";
                    gridTextColor = alpha + "000000";
                    watermarkTextColor = alpha + "8b8b8b";
			    } else {
                    bgColor = alpha + "FFFFFF";
                    lineColor =alpha +  "000000";
                    headerTextColor = alpha + "000000";
                    scaleOneColor =alpha +  "FFFFFF";
                    scaleTwoColor = alpha + "FFFFFF";
                    gridTextColor = alpha + "000000";
                    watermarkTextColor =alpha + "000000";
			    }
		    }
	    }
	
	    public void setWatermark(String mark) {
		    watermark = mark;	
	    }
    }
}
