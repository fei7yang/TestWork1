package com.dfl.report.mfcadd;


import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;

import javax.imageio.ImageIO;

import org.apache.poi.hslf.model.Comment;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class POIExcel
{
	Workbook wb = null;
	Sheet sheet = null;

	FileInputStream fis = null;
	int type = 0;//xlsx
	InputStream is = null;
	public void init() {
        wb = new XSSFWorkbook();
        sheet = wb.getSheetAt(0);
    }
	public void specifyTemplate(InputStream fisin) throws FileNotFoundException,
    IOException {
		is = fisin;
		if(wb==null)
			wb = new XSSFWorkbook(fisin);
		sheet =  wb.getSheetAt(0);
}
	/**
	 * @return the sheet
	 */
	public Sheet getSheet() {
		return sheet;
	}

	/**
	 * @param sheet the sheet to set
	 */
	public void setSheet(Sheet sheet) {
		this.sheet = sheet;
	}
	public void specifyTemplate(String filePath) throws FileNotFoundException, IOException
	{
		if (wb == null)
		{
			fis = new FileInputStream(filePath);
			if(filePath.toLowerCase().endsWith(".xls")){
				wb = new HSSFWorkbook(fis);
				type = 1;
			}else{
				 wb = new XSSFWorkbook(fis);
				 type = 0;
			}
		}
		sheet = wb.getSheetAt(0);
//		try{
//			int numbers = wb.getNumberOfNames();
//	    	for(int i = 0 ; i < numbers; i ++){
//	    		Name name = wb.getNameAt(i);
//	    		System.out.println("name.getRefersToFormula() := " + name.getRefersToFormula());
//	    		System.out.println("name.getNameName() := " + name.getNameName());
//	    	}
//		}catch(Exception e){
//			e.printStackTrace();
//		}
	}
	public void removeSheet(int index){
		try{
			wb.removeSheetAt(index);
		}catch(Exception e){
			e.printStackTrace();
		}
	}
	public void specifyTemplate(String filePath, int sheetID)// throws FileNotFoundException, IOException
	{
		try{
			fis = new FileInputStream(filePath);
			if(filePath.toLowerCase().endsWith(".xls")){
				wb = new HSSFWorkbook(fis);
				type = 1;
			}else if(filePath.toLowerCase().endsWith(".xlsx")){
				 wb = new XSSFWorkbook(fis);
				 type = 0;
			}
			sheet = wb.getSheetAt(sheetID);
		}catch(Exception e){
			e.printStackTrace();
		}
	}
	public void specifyTemplate(String filePath, String sheetName)// throws FileNotFoundException, IOException
	{
		try{
			if(wb == null){
				fis = new FileInputStream(filePath);
				if(filePath.toLowerCase().endsWith(".xls")){
					wb = new HSSFWorkbook(fis);
					type = 1;
				}else if(filePath.toLowerCase().endsWith(".xlsx")){
					 wb = new XSSFWorkbook(fis);
					 type = 0;
				}
			}
			
			sheet = wb.getSheet(sheetName);
		}catch(Exception e){
			e.printStackTrace();
		}
	}
	public void zoomSheet(int sheetIndex, int rows, int columns, int molecule, boolean landscape ) {
		//sheet.setZoom(molecule, molecule);
		   wb.setPrintArea(sheetIndex, 0, columns, 0, rows);
		   PrintSetup printSetup = sheet.getPrintSetup();
		   printSetup.setPaperSize(PrintSetup.A4_PAPERSIZE);
		   printSetup.setScale((short) molecule);// 自定义缩放，此处100为无缩放
		   printSetup.setLandscape(landscape); // 打印方向，true：横向，false：纵向(默认)
	}
	public void specifyTemplate(int sheetID) throws FileNotFoundException, IOException
	{
		if (wb != null)
		{
			sheet = wb.getSheetAt(sheetID);
		}
		
	}
	public void outputExcel(String outPath, boolean flag) throws FileNotFoundException, IOException
	{
		if(flag) {
			try{
				sheet.setForceFormulaRecalculation(true);
				wb.setForceFormulaRecalculation(true);
			}catch(Exception e) {
				
			}
		}
		
		FileOutputStream fileOut = new FileOutputStream(outPath); // "c:/workbook.xls"
		wb.write(fileOut);
		fileOut.close();
		if(fis != null) {
			fis.close();
		}
		if(is != null) {
			is.close();
		}
//		File xlsFile = new File(outPath);
//		if(xlsFile.exists()){
//			String pdfPath = outPath.replace(".xlsx", ".pdf");
//			if(type == 1){
//				pdfPath = outPath.replace(".xls", ".pdf");
//			}
//			saveExcelAsPdf(outPath, pdfPath);
//		}
	}
	public void outputExcel(String outPath) throws FileNotFoundException, IOException
	{
		
		FileOutputStream fileOut = new FileOutputStream(outPath); // "c:/workbook.xls"
		wb.write(fileOut);
		fileOut.close();
		if(fis != null) {
			fis.close();
		}
		if(is != null) {
			is.close();
		}
//		File xlsFile = new File(outPath);
//		if(xlsFile.exists()){
//			String pdfPath = outPath.replace(".xlsx", ".pdf");
//			if(type == 1){
//				pdfPath = outPath.replace(".xls", ".pdf");
//			}
//			saveExcelAsPdf(outPath, pdfPath);
//		}
	}
	/**
     * 20100509克隆模板
//     */
//    public void cloneTemplate(String[] names){
//    	for(int i = 0; i < names.length - 1; i ++){
//    		wb.cloneSheet(0);
//    	}
//    	for(int i = 0 ; i < names.length ; i ++){
//    		wb.setSheetName(i, names[i]);
//    	}
//    }
	public void cloneTemplate(String[] names){
    	for(int i = 0; i < names.length - 1; i ++){
    		wb.cloneSheet(0);
    	}
    	for(int i = 0 ; i < names.length ; i ++){
    		wb.setSheetName(i, names[i]);
    	}
    	int numbers = wb.getNumberOfNames();
    	for(int i = 0 ; i < numbers; i ++){
    		Name name = wb.getNameAt(i);
    		for(int j = 1; j < names.length ; j ++){
    			Name newName = wb.createName();
    			newName.setSheetIndex(j);
    			String newRefers = name.getRefersToFormula().replaceAll(names[0], names[j]);
    			//System.out.println("newRefers := " + newRefers);
    			newName.setRefersToFormula(newRefers);
    			newName.setNameName(name.getNameName());
    		}
    	}
    }
    public void cloneTemplate(int index, String[] names){
    	for(int i = 0; i < names.length - 1; i ++){
    		wb.cloneSheet(index);
    	}
    	for(int i = 0 ; i < names.length ; i ++){
    		wb.setSheetName(index + i, names[i]);
    	}
    }
    public void renameSheet(int index, String name){
    	wb.setSheetName(index, name);
    }
    public static boolean copyFile (String output ,String input)throws FileNotFoundException, IOException {
       
            
		try
		{
			 File fl = new File(output);
		        if (fl.exists()){
		        	fl.delete();
		        }
			StringBuffer cmd = new StringBuffer();
			cmd.append("cmd  /c  copy  \"");
			cmd.append(input);
			cmd.append("\" \"");
			cmd.append(output);
			cmd.append("\"");
			Process process = Runtime.getRuntime().exec(cmd.toString());
			process.waitFor();
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
        return true;
    }

	public static void openExcel(String wbPath)
	{

		try
		{
			StringBuffer cmd = new StringBuffer();
			cmd.append("cmd  /c  start \"");
			cmd.append(wbPath);
			cmd.append("\"");
			Runtime.getRuntime().exec(cmd.toString());
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
	}
	public void fillCellValue(int row, int col, String cellValue) {
		if(cellValue == null){
			cellValue = "";
		}
		try{
			if(type == 0){
				XSSFRow xfRow = ((XSSFSheet)sheet).getRow(row);
				if(xfRow == null)System.out.println("row := " + row);
				XSSFCell xfCell = xfRow.getCell(col);
				//xfCell.set
				xfCell.setCellValue(cellValue);
			}else if(type == 1){
				HSSFRow hfRow = ((HSSFSheet)sheet).getRow(row);
				HSSFCell hfCell = hfRow.getCell(col);
				hfCell.setCellValue(cellValue);
			}
		}catch(Exception e){
			e.printStackTrace();
		}
    }
	public void deleteRows(int startRow, int endRow){
//		for(int i = endRow; i >= startRow ; i --){
//			 Row row = sheet.getRow(i);
//			 if(row != null){
//				 sheet.removeRow(row);
//			 }
//		} 
		List<CellRangeAddress> lstMerges = new ArrayList<CellRangeAddress>();
		
		for(int i = 0; i < sheet.getNumMergedRegions(); i++){
	    	CellRangeAddress mergedold =  sheet.getMergedRegion(i);
	    	 boolean inStart = mergedold.getFirstRow() >= startRow || mergedold.getLastRow() >= startRow;
	            boolean inEnd = mergedold.getLastRow() <= endRow || mergedold.getFirstRow() <= endRow;
	            if (inStart && inEnd){
	            	lstMerges.add(mergedold);
	            	//sheet.removeMergedRegion(i);
	            }
	    }
		for(int  i = 0 ; i < lstMerges.size() ; i ++){
			CellRangeAddress mergedold =  lstMerges.get(i);
			for(int j = 0; j < sheet.getNumMergedRegions(); j ++){
				CellRangeAddress merge =  sheet.getMergedRegion(j);
				if(merge == mergedold){
					sheet.removeMergedRegion(j);
				}
			}
		}
	    //for(int i)
	    sheet.createRow(endRow + 1);
	    for(int i = 0 ; i < (endRow - startRow + 1); i ++){
	    	sheet.shiftRows(endRow + 1 - i, endRow + 2 - i, -1);
	    }
	}
	
	public void copyTemplate( int times,int startRow, int endRow, int n, boolean copyRowHeight){
        int s;
        int inc;
        if (n < 0) {
            s = startRow;
            int e = endRow;
            inc = 1;
        } else {
            s = endRow;
            int e = startRow;
            inc = -1;
        }
        ArrayList<Name> alNames = new ArrayList<Name>();
        int numbers = wb.getNumberOfNames();
        for(int i = 0 ; i < numbers; i ++){
        	Name name = wb.getNameAt(i);
        	String formula = name.getRefersToFormula();
        	String[] splits = formula.split("\\$");
        	try{
        		int row = Integer.parseInt(splits[splits.length - 1]);
        		if(row >= startRow + 1 && row <= endRow + 1){
        			alNames.add(name);
        		}
        	}catch(Exception e){
        		
        	}
        }
//        int numbers = wb.getNumberOfNames();
//    	for(int i = 0 ; i < numbers; i ++){
//    		Name name = wb.getNameAt(i);
//    		for(int j = 1; j < names.length ; j ++){
//    			Name newName = wb.createName();
//    			newName.setSheetIndex(j);
//    			String newRefers = name.getRefersToFormula().replaceAll(names[0], names[j]);
//    			//System.out.println("newRefers := " + newRefers);
//    			newName.setRefersToFormula(newRefers);
//    			newName.setNameName(name.getNameName());
//    		}
//    	}
        for (int rowNum = s; rowNum >= startRow && rowNum <= endRow && rowNum >= 0 &&
                          rowNum < 0x10000; rowNum += inc) {
            Row row = sheet.getRow(rowNum);
            if (row == null)
                    continue;
            for(int i = 1 ; i<= times ; i++){
            	//System.out.println("rowNum + n*i " + (rowNum + n*i));
            	Row row2Replace = sheet.getRow(rowNum + n*i);
                if (row2Replace == null)
                    row2Replace = sheet.createRow(rowNum + n*i);
               // System.out.println("row.getFirstCellNum() := " + (row.getFirstCellNum()));
               // System.out.println("row.getLastCellNum() := " + (row.getLastCellNum()));
                for (int col = row2Replace.getFirstCellNum();
                                 col <= row2Replace.getLastCellNum(); col++) {
                	if(col < 0){
                		continue;
                	}
                    Cell cell = row2Replace.getCell(col);
                    if (cell != null)
                        row2Replace.removeCell(cell);
                }
                if (copyRowHeight)
                    row2Replace.setHeight(row.getHeight());
                for (int col = row.getFirstCellNum();
                                 col <= row.getLastCellNum();
                                 col++) {
                    Cell cell = row.getCell(col);
                    if (cell != null) {
                    	Cell newCell = row2Replace.createCell(col);
                    	copyCell(cell, newCell, true);
                    	//newCell.setCellValue(cell.getStringCellValue());
                    }
                }
                
            }
        }
        
        if(alNames.size() > 0){
        	for(int x = 0; x < alNames.size() ; x ++){
        		Name name = alNames.get(x);
            	String formula = name.getRefersToFormula();
            	String[] splits = formula.split("\\$");
            	try{
            		int oldNameRow = Integer.parseInt(splits[splits.length - 1]);
            		for(int i = 1 ; i<= times ; i++){
            			int newNameRow = oldNameRow + n * i;
            			String newRefers = formula.replaceAll(splits[splits.length - 1], newNameRow + "");
            			Name newName = wb.createName();
            			newName.setRefersToFormula(newRefers);
            			newName.setNameName(name.getNameName() + i);
            		}
            	}catch(Exception e){
            		
            	}
        	}
			
        }
    }
	public void addMergedRegion(int startRow, int startCell, int endRow, int endCell){
		CellRangeAddress merged = new CellRangeAddress(startRow, endRow,  startCell, endCell);
		sheet.addMergedRegion(merged );
	}
	public void addBackgroundColor(int row, int col, int r, int g , int b){
		Row rowOld = sheet.getRow(row);
		if (rowOld != null) {
			Cell cellOld = rowOld.getCell(col);
			if (cellOld != null) {
				CellStyle style = wb.createCellStyle();
				if(type == 0){
					((XSSFCellStyle)style).setFillForegroundColor(new XSSFColor(new java.awt.Color(r,g,b)));
				}else{
					HSSFPalette palette = ((HSSFWorkbook)wb).getCustomPalette();
					palette.setColorAtIndex(HSSFColor.RED.index, (byte)r, (byte)g, (byte)b);
					style.setFillForegroundColor(HSSFColor.RED.index);
				}
				style.setFillPattern(CellStyle.SOLID_FOREGROUND);
				style.setBorderBottom(CellStyle.BORDER_THIN); // 下边框
				style.setBorderLeft(CellStyle.BORDER_THIN);// 左边框
				style.setBorderTop(CellStyle.BORDER_THIN);// 上边框
				style.setBorderRight(CellStyle.BORDER_THIN);// 右边框
				style.setAlignment(CellStyle.ALIGN_LEFT);
				style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
				cellOld.setCellStyle(style);
			}
		}
	}
	public void addBackgroundColor(int row, int col, short color) {
		Row rowOld = sheet.getRow(row);
		if (rowOld != null) {
			Cell cellOld = rowOld.getCell(col);
			if (cellOld != null) {
				CellStyle style = wb.createCellStyle();
				style.setFillForegroundColor(color);
				style.setFillPattern(CellStyle.SOLID_FOREGROUND);
				style.setBorderBottom(CellStyle.BORDER_THIN); // 下边框
				style.setBorderLeft(CellStyle.BORDER_THIN);// 左边框
				style.setBorderTop(CellStyle.BORDER_THIN);// 上边框
				style.setBorderRight(CellStyle.BORDER_THIN);// 右边框
				style.setAlignment(CellStyle.ALIGN_CENTER);
				style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
				cellOld.setCellStyle(style);
			}
		}
	}
	public void addForeColor(int row, int col, short color){
		Row rowOld = sheet.getRow(row);
		 if(rowOld != null){
			 Cell cellOld = rowOld.getCell(col);
			 if(cellOld != null){
				CellStyle styleNew = wb.createCellStyle();
				// style.setFillForegroundColor(color);
					//style.setFillPattern(CellStyle.SOLID_FOREGROUND);
				styleNew.setBorderBottom(CellStyle.BORDER_THIN); //下边框    
				styleNew.setBorderLeft(CellStyle.BORDER_THIN);//左边框    
				styleNew.setBorderTop(CellStyle.BORDER_THIN);//上边框    
				styleNew.setBorderRight(CellStyle.BORDER_THIN);//右边框
				styleNew.setAlignment(CellStyle.ALIGN_CENTER);
				styleNew.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
				Font newFont = this.cloneFont(wb.getFontAt(cellOld.getCellStyle().getFontIndex()));
				newFont.setColor(color);
//				 RichTextString rts = cellOld.getRichStringCellValue();
//				 rts.applyFont(newFont);
//				 cellOld.setCellValue(rts);
				styleNew.setFont(newFont);
				cellOld.setCellStyle(styleNew);
			 }
		 }
	}
	private Font cloneFont(Font font){
		Font newFont = wb.createFont();
		newFont.setBoldweight(font.getBoldweight());
		newFont.setColor(font.getColor());
		newFont.setFontHeightInPoints(font.getFontHeightInPoints());
		newFont.setFontName(font.getFontName());
		newFont.setItalic(font.getItalic());
		newFont.setStrikeout(font.getStrikeout());
		newFont.setTypeOffset(font.getTypeOffset());
		newFont.setUnderline(font.getUnderline());
		return newFont;
	}
	public void copyMergedRegion(int times , int startRow, int endRow, int n){
	    List shiftedRegions = new ArrayList();
	    List mergedolds = new ArrayList();
	    for(int i = 0; i < sheet.getNumMergedRegions(); i++){
	        mergedolds.add( sheet.getMergedRegion(i));
	    }
	    for(int i = 0; i < mergedolds.size(); i++)
	    {
	        CellRangeAddress mergedold = (CellRangeAddress)mergedolds.get(i);
	        for (int p = 1 ; p<= times ; p++){
	        	CellRangeAddress merged = new CellRangeAddress(mergedold.getFirstRow(), mergedold.getLastRow(), 
	        			mergedold.getFirstColumn(), mergedold.getLastColumn());
	        	
	            boolean inStart = merged.getFirstRow() >= startRow ||
	                              merged.getLastRow() >= startRow;
	            boolean inEnd = merged.getLastRow() <= endRow ||
	                            merged.getFirstRow() <= endRow;
	            if (inStart && inEnd && !merged.isInRange(startRow - 1, 0) &&
	                !merged.isInRange(endRow + 1, 0)) {
	                merged.setFirstRow(merged.getFirstRow() + n*p);
	                merged.setLastRow(merged.getLastRow() + n*p);
	                shiftedRegions.add(merged);
	            }
	        }
	    }
	    CellRangeAddress region;
	    for(Iterator iterator = shiftedRegions.iterator(); iterator.hasNext(); sheet.addMergedRegion(region))
	        region = (CellRangeAddress)iterator.next();

	}
	public void copyRow(Sheet sheet, Row oldRow, Row newRow, int[] rowlocator, String[] newrowDats)
 	{
 		//Set mergedRegions = new HashSet();
 		if (oldRow.getHeight() >= 0)
 			newRow.setHeight(oldRow.getHeight());
 		for (int j = oldRow.getFirstCellNum(); j <= oldRow.getLastCellNum(); j++)
 		{
 			Cell oldCell = oldRow.getCell((short)j);
 			Cell newCell = newRow.getCell((short)j);
 			if (oldCell == null)
 				continue;
 			if (newCell == null)
 				newCell = newRow.createCell((short)j);
 			copyCell(oldCell, newCell, true);
 		}
 		for(int i = 0 ; i < rowlocator.length ; i ++){
 			Cell newCell = newRow.getCell((short)rowlocator[i]);
 			newCell.setCellValue(newrowDats[i]);
 		}
 	}
	public void copyRow(Sheet sheet, Row oldRow, Row newRow)
 	{
 		//Set mergedRegions = new HashSet();
 		if (oldRow.getHeight() >= 0)
 			newRow.setHeight(oldRow.getHeight());
 		for (int j = oldRow.getFirstCellNum(); j <= oldRow.getLastCellNum(); j++)
 		{
 			Cell oldCell = oldRow.getCell((short)j);
 			Cell newCell = newRow.getCell((short)j);
 			if (oldCell == null)
 				continue;
 			if (newCell == null)
 				newCell = newRow.createCell((short)j);
 			copyCell(oldCell, newCell, true);
 		}
 	}
	 public void appendRow(int startRow,int rows) {   
         
    	 
          //sheet.shiftRows(startRow + 1, sheet.getLastRowNum(), rows,true,false);   
          
  
          for (int i = 0; i < rows; i++) {   
                 
                Row sourceRow = null;   
                Row targetRow = null;   
                   
                sourceRow = sheet.getRow(startRow);   
                targetRow = sheet.createRow(++startRow);   
                   
                copyRow(sheet, sourceRow, targetRow);   
          }   
             
    } 
	 public void insertRow(int startRow,int rows) {   
         
    	 
         sheet.shiftRows(startRow + 1, sheet.getLastRowNum(), rows,true,false);   
         
 
         for (int i = 0; i < rows; i++) {   
                
               Row sourceRow = null;   
               Row targetRow = null;   
                  
               sourceRow = sheet.getRow(startRow);   
               targetRow = sheet.createRow(++startRow);   
                  
               copyRow(sheet, sourceRow, targetRow);   
         }   
            
   } 
	 public void copyCell(int oldRow, int oldCell, int newRow, int startCell, int endCell, boolean copyStyle){
		 Row rowOld = sheet.getRow(oldRow);
		 if(rowOld != null){
			 Cell cellOld = rowOld.getCell(oldCell);
			 if(cellOld != null){
				 Row rowNew = sheet.getRow(newRow);
				 if(rowNew == null){
					 rowNew = sheet.createRow(newRow);
				 }
				 int width = sheet.getColumnWidth(oldCell);
				 for(int i = startCell; i <= endCell; i++){
					 Cell cellNew = rowNew.getCell(i);
					 if(cellNew == null){
						 cellNew = rowNew.createCell(i);
					 }
					 copyCell(cellOld, cellNew, copyStyle);
					
					 if(copyStyle) {
						 sheet.setColumnWidth(i, width);
					 }
				 }
			 }
		 }
	 }
	 public void copyCell(Cell oldCell, Cell newCell, boolean copyStyle)
	 	{
	 		if (copyStyle) {
	 			//newCell.setCellStyle(oldCell.getCellStyle());
	 			CellStyle oldStyle = oldCell.getCellStyle();
	 			CellStyle newStyle = oldCell.getCellStyle();
	 			newStyle.cloneStyleFrom(oldStyle);
	 			newCell.setCellStyle(newStyle);
	 		}
	 		//newCell.setEncoding(HSSFCell.ENCODING_UTF_16);
	 		switch (oldCell.getCellType())
	 		{
	 		case 1: // '\001'
	 			newCell.setCellValue(oldCell.getStringCellValue());
	 			break;

	 		case 0: // '\0'
	 			newCell.setCellValue(oldCell.getNumericCellValue());
	 			break;

	 		case 3: // '\003'
	 			newCell.setCellType(3);
	 			break;

	 		case 4: // '\004'
	 			newCell.setCellValue(oldCell.getBooleanCellValue());
	 			break;

	 		case 5: // '\005'
	 			newCell.setCellErrorValue(oldCell.getErrorCellValue());
	 			break;

	 		case 2: // '\002'
	 			newCell.setCellFormula(oldCell.getCellFormula());
	 			break;
	 		}
	 	}
	 float getRowHeightPixel(int row){
		 return sheet.getRow(row).getHeightInPoints();
	 }
	 public void autoSizeColumn(int col){
		 sheet.autoSizeColumn(col, true);
	 }
	 public void addComment(int row, int col, String author, String comment){
		Drawing par = sheet.createDrawingPatriarch();
		Row r = sheet.getRow(row);
		Cell c = r.getCell(col);
		//System.out.println("author := " + author);
		 if(type == 0){
		    	XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 1023, 255, col, row, col + 2, row + 3);
		    	org.apache.poi.ss.usermodel.Comment com = par.createCellComment(anchor);
		    	com.setString(new XSSFRichTextString(comment));
		    	com.setAuthor(author);
		    	c.setCellComment(com);
		    }else{
		    	HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 0, 0, (short)col, row, (short)col, row );
			    //drawing.createPicture(anchor, pictureIdx);
		    	org.apache.poi.ss.usermodel.Comment com = par.createCellComment(anchor);
		    	com.setString(new HSSFRichTextString(comment));
		    	com.setAuthor(author);
		    	c.setCellComment(com);
		    }
	 }
	 /**
		 * @param col1 列1
		 * @param col2 列2
		 * @param row1 行
		 * @param dx1 图片偏移像素数
		 * @param bufImage 图片
		 */
		public void writeImage(int col1, int col2, int row1, int row2, BufferedImage bufImage)
		{
			//int x1 = dx1;
			try {
				ImageIO.setUseCache(false);
				ByteArrayOutputStream baos = new ByteArrayOutputStream();
				ImageIO.write(bufImage, "jpg", baos);
				byte[] bytes = baos.toByteArray();
				int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
				// 创建绘图族，  这是所有形状的顶级容器 
			   Drawing drawing = sheet.createDrawingPatriarch();
			    if(type == 0){
			    	XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 1023, 255, col1, row1, col2, row2);
				    drawing.createPicture(anchor, pictureIdx);
			    }else{
			    	HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 1023, 255, (short)col1, row1, (short)(col2 - 1), row2 - 1);
				    drawing.createPicture(anchor, pictureIdx);
			    }
			    
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		/**
		 * 读取 Excel文件的信息 保存到两维数组 中?
		 * 
		 * @return
		 */
		public String[][] readinfo() {
			Row row = null;
			Cell cell = null;
			int i = 0;
			int rowCount = sheet.getLastRowNum() +1;
			int columns = 0;
			for(i = 0 ; i < rowCount ; i ++){
				Row r = sheet.getRow(i);
				if(r != null){
					int col = r.getLastCellNum();
					//System.out.println("col = "+col);
					if(columns < col){
						columns = col;
					}
				}
			}
			String[][] infos = new String[rowCount][columns];
			ArrayList<String[]> alDatas = new ArrayList<String[]>();
			//for (Iterator it = sheet.rowIterator(); it.hasNext() && i <= rowCount; i++) {
			for(i = 0 ; i < rowCount ; i ++){
				//row = (Row) it.next();
				row = sheet.getRow(i);
				if(row != null){
					boolean allBlank = true;
					for(int j = 0 ; j < columns; j ++){
						cell = row.getCell(j);
						if(cell == null){
							infos[i][j] = "";
						}else if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC){
							NumberFormat nf = NumberFormat.getInstance(Locale.SIMPLIFIED_CHINESE);
							nf.setMinimumFractionDigits(0);
							nf.setGroupingUsed(false);
							infos[i][j] = nf.format(cell.getNumericCellValue());
						}else{
							infos[i][j] = cell.getStringCellValue().trim();
						}
						if(infos[i][j].length() > 0){
							allBlank = false;
						}
					}
					if(allBlank){
						break;
					}
					alDatas.add(infos[i]);
				}
			}
			infos = new String[alDatas.size()][columns];
			for(i = 0; i < alDatas.size(); i ++){
				infos[i] = alDatas.get(i);
			}
			return infos;
		}
		public void close(){
			if(this.fis != null){
				try{
					fis.close();
				}catch(Exception e){
					e.printStackTrace();
				}finally{
					fis = null;
				}
			}
		}
		public void groupRow(int startRow, int endRow){
			sheet.groupRow(startRow, endRow);
		}
		public void groupColumn(int startCol, int endCol){
			sheet.groupColumn(startCol, endCol);
		}
		public int getSheetRows(){
			return sheet.getLastRowNum() + 1;
		}
	public static void main(String[] args) throws FileNotFoundException, IOException, ClassNotFoundException
	{
//		System.out.println("69   23   23".trim().split(" ").length);
//		testing();
		POIExcel poi = new POIExcel();
		poi.specifyTemplate("e:\\1.xlsx",0);
		poi.zoomSheet(0, 13, 32, 50, true);
		//poi.fillCellValue(5, 3, "    工厂工程：工厂焊装工程");
		//poi.fillCellValue(6, 3, "    车    型：" + "");
//		poi.fillCellValue(7, 3, "    版    次：" + "A");
//		String docNo = "" + "-" + "" ;
//		poi.fillCellValue(8, 3, "    文件编号：" + docNo+ "-GG-A");
//		poi.fillCellValue(10, 3, "    编制日期：" + "");
//		poi.groupRow(2, 12);
//		poi.groupRow(3, 4);
//		//poi.groupRow(3, 4);
//		poi.groupRow(6, 10);
//		poi.groupRow(12, 12);
//		poi.groupRow(9, 10);
//		poi.groupRow(2, 4);
//		poi.groupRow(8, 10);
//		poi.groupRow(5, 10);
//		poi.groupRow(11, 12);
	//	System.out.println(poi.getSheetRows());
//		poi.copyTemplate(3, 47, 93, 47, true);
//		poi.copyMergedRegion(3, 47, 93, 47);
		//poi.insertRow(39, 10);
		//poi.removeSheet(1);
		//poi.insertRow(2, 3);
//		poi.groupRow(2, 8);
//		poi.groupRow(3, 6);
//		poi.groupRow(5, 6);
////		poi.groupRow(2, 7);
//		
//		
//		poi.groupRow(11, 14);
//		poi.groupRow(41, 55);
//		poi.groupRow(56, 56);
//		poi.groupRow(57, 57);
		//poi.insertRow(32, 10);
		//poi.copyCell(1, 3, 1, 4, 50, true);
		//poi.copyCell(2, 3, 2, 4, 50, true);
		//poi.addMergedRegion(0, 0, 0, 50);
		//sheet.addMergedRegion(new Region(0, (short)0, 0, (short)maxColumnNum));
		//poi.addComment(2, 2, "规划部基建计划模板\r\n什么");
		//poi.addForeColor(2, 2, IndexedColors.GREEN.getIndex());
//		poi.addBackgroundColor(12, 10, IndexedColors.INDIGO.getIndex());
//		poi.addBackgroundColor(13, 10, IndexedColors.VIOLET.getIndex());
//		poi.addBackgroundColor(14, 10, IndexedColors.TURQUOISE.getIndex());
//		poi.addBackgroundColor(15, 10, IndexedColors.TEAL.getIndex());
//		poi.addBackgroundColor(16, 10, IndexedColors.TAN.getIndex());
//		poi.addBackgroundColor(17, 10, IndexedColors.SEA_GREEN.getIndex());
//		poi.addBackgroundColor(18, 10, IndexedColors.SKY_BLUE.getIndex());
//		poi.addBackgroundColor(19, 10, IndexedColors.ROYAL_BLUE.getIndex());
//		poi.addBackgroundColor(20, 10, IndexedColors.OLIVE_GREEN.getIndex());
//		poi.appendRow(2, 5);
//		int newCols = 5;
//		if(newCols > 1) {
//			poi.copyCell(0, 7, 0, 8, 6 + newCols, true);
//			poi.copyCell(1, 7, 1, 8, 6 + newCols, true);
//			poi.copyCell(2, 7, 2, 8, 6 + newCols, true);
//			for(int row = 3; row < 11; row ++) {
//				poi.copyCell(3, 7, row, 8, 6 + newCols, true);
//			}
//			poi.addMergedRegion(0, 7, 0, 6 + newCols);
//			poi.addMergedRegion(1, 7, 1, 6 + newCols);
//			for(int col = 0; col < newCols; col ++) {
//				poi.fillCellValue(2, 7 + col, "○");
//			}
//		}
		poi.outputExcel("e:\\123.xlsx");
		
		
		poi.close();
//		poi.copyTemplate(3, 0, 37, 37, true	);
//		poi.copyMergedRegion(3, 0, 37, 37);
		//poi.writeImage(15, 31, 6, 32, ImageIO.read(new File("C:\\123.jpg")));
		//poi.outputExcel("C:\\2.xls");
//		POIExcel.copyFile("E:\\mifc.xlsx", "E:\\123.xlsx");
//		//poi.specifyTemplate("E:\\mifc.xlsx");
//		POIExcel.openExcel("E:\\mifc.xlsx");
	}
}
