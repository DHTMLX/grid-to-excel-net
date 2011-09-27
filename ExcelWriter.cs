using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

using System.Web;

using OpenExcel.OfficeOpenXml.Style;
using OpenExcel.OfficeOpenXml;
//using DocumentFormat.OpenXml.Spreadsheet;
///TODO: add text color, vertical alignment, horizontal alignment, colspan, rowspan
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

	    private String watermark = null;
	
	
        public void generate(string xml, Stream output){
            parser = new ExcelXmlParser();
            //    try {
            parser.setXML(xml);
            createExcel(output);
            setColorProfile();
            headerPrint(parser);
          
            rowsPrint(parser, output);
            wb.Workbook.Document.Styles.Save();
            footerPrint(parser);
            insertHeader(parser, output);
            insertFooter(parser, output);
            watermarkPrint(parser);
         
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

	    }

        public void generate(HttpContext context)
        {
            generate(context.Server.UrlDecode(context.Request.Form["grid_xml"]), context.Response);
        }

        public void generate(string xml, HttpResponse resp)
        {
            var data = new MemoryStream();
            
            resp.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
		    resp.HeaderEncoding = Encoding.UTF8;
		    resp.AppendHeader("Content-Disposition", "attachment;filename=grid.xls");
		    resp.AppendHeader("Cache-Control", "max-age=0");
            generate(xml, data);
            data.Close();
            data.WriteTo(resp.OutputStream);
            
            
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
                if (headerTextColor != "FF000000")
                    font.Color = headerTextColor;

                ExcelBorder border = getBorder();  


			    for (uint row = 1; row <= cols.Length; row++) {
                    
                    sheet.Rows[row].Height = 22.5;
				    for (uint col = 1; col <= cols[row-1].Length; col++) {

                        
                        sheet.Cells[row, col].Style.Font = font;//if bold font assigned after border - all table will be bold, weird, find out later
                        
                        sheet.Cells[row, col].Style.Border = border;

                        sheet.Columns[col].Width = widths[col-1] / scale;
					    String name = cols[row-1][col-1].GetName();
                        if (bgColor != "FFFFFFFF")
                            sheet.Cells[row, col].Style.Fill.ForegroundColor = bgColor;
                        

                        

                        ///TODO: 
                        ///font color, merge cells, alignment
                        sheet.Cells[row, col].Value = name;
					    colsNumber = (int)col;
				    }
			    }
			    headerOffset = cols.Length;
                sheet.MergeTwoCells("A1", "B1");


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


        protected ExcelBorder getBorder()
        {
            ExcelBorder border = new ExcelBorder(null, wb.Styles, 0);           
            border.BottomStyle = OpenExcel.OfficeOpenXml.Style.ExcelBorderStyleValues.Thin;
            border.BottomColor = lineColor;
            border.LeftStyle = OpenExcel.OfficeOpenXml.Style.ExcelBorderStyleValues.Thin;
            border.LeftColor = lineColor;
            border.RightStyle = OpenExcel.OfficeOpenXml.Style.ExcelBorderStyleValues.Thin;
            border.RightColor = lineColor;
            border.TopStyle = OpenExcel.OfficeOpenXml.Style.ExcelBorderStyleValues.Thin;
            border.TopColor = lineColor;
            return border;
        }

	    private void footerPrint(ExcelXmlParser parser){
		    cols = parser.getColumnsInfo("foot");
            ExcelBorder border = getBorder();  
		    if (parser.getWithoutHeader() == false) {
                ExcelFont font = wb.CreateFont("Arial", 10);
                
                font.Bold = true;
                if (headerTextColor != "FF000000")
                    font.Color = headerTextColor;
			    for (uint row = 1; row <= cols.Length; row++) {
                    
                    uint rowInd = (uint)(row + headerOffset);
                    sheet.Rows[rowInd].Height = 22.5;
                 
				    for (uint col = 1; col <= cols[row-1].Length; col++) {
					
                        if (bgColor != "FFFFFFFF")
                            sheet.Cells[rowInd, col].Style.Fill.ForegroundColor = bgColor;
                        sheet.Cells[rowInd, col].Style.Font = font;
                        //TODO add text color, vertical alignment, horizontal alignment
                        sheet.Cells[rowInd, col].Style.Border = border;
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
            ExcelFont font = wb.CreateFont("Arial", 10);
            font.Bold = true;
            font.Color = watermarkTextColor;
            
		    ExcelBorder border = getBorder();

		   // f.setAlignment(Alignment.CENTRE);
            sheet.Cells[(uint)(headerOffset + 1), 0].Value = watermark;
		  //  Label label = new Label(0, headerOffset, watermark , f);
		  //  sheet.addCell(label);
		   // sheet.mergeCells(0, headerOffset, colsNumber, headerOffset);*/
	    }

	    private void rowsPrint(ExcelXmlParser parser, Stream resp) {
		    
		    ExcelRow[] rows = parser.getGridContent();

		    this.rows_stat = rows.Length;
           
            ExcelBorder border = getBorder();
            ExcelFont font = wb.CreateFont("Arial", 10);
           // if (gridTextColor != "FF000000")
           //      font.Color = gridTextColor;

		    for (uint row = 1; row <= rows.Length; row++) {
			    ExcelCell[] cells = rows[row-1].getCells();
                uint rowInd = (uint)(row + headerOffset);
                sheet.Rows[rowInd].Height = 20;
	 
			    for (uint col = 1; col <= cells.Length; col++) {




                    if (cells[col - 1].GetBold() || cells[col - 1].GetItalic())
                    {
                        ExcelFont curFont = wb.CreateFont("Arial", 10); ;
                       // if (gridTextColor != "FF000000")
                     //       font.Color = gridTextColor;
                        if (cells[col - 1].GetBold())
                            font.Bold = true;

                        if (cells[col - 1].GetItalic())
                            font.Italic = true;

                        sheet.Cells[rowInd, col].Style.Font = curFont;
                    }
                    else
                    {
                        sheet.Cells[rowInd, col].Style.Font = font;
                    }

                    sheet.Cells[rowInd, col].Style.Border = border;
   

                    if ((!cells[col - 1].GetBgColor().Equals(""))&&(parser.getProfile().Equals("full_color"))) {
                        sheet.Cells[rowInd, col].Style.Fill.ForegroundColor = "FF" + cells[col - 1].GetBgColor();
				    } else {
					    //Colour bg;
                        if (row % 2 == 0 && scaleTwoColor != "FFFFFFFF")
                        {
                            sheet.Cells[rowInd, col].Style.Fill.ForegroundColor = scaleTwoColor;						
					    } else {
                            if (scaleOneColor != "FFFFFFFF")
                                sheet.Cells[rowInd, col].Style.Fill.ForegroundColor = scaleOneColor;
					    }
				    }

                    
                    int intVal;
                    double dbVal;

                    if (int.TryParse(cells[col - 1].GetValue(), out intVal))
                    {
                        sheet.Cells[rowInd, col].Value = intVal;
                    }
                    else if (double.TryParse(cells[col - 1].GetValue(), out dbVal))
                    {
                        sheet.Cells[rowInd, col].Value = dbVal;
                    }
                    else
                    {
                        sheet.Cells[rowInd, col].Value = cells[col - 1].GetValue();
                    }
                        
                    
                    //COLOR!
				   
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
