package com.dfl.report.util;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.swt.widgets.TableItem;

public class OutputDataToExcel1 {
	private XSSFWorkbook book;
	private XSSFCellStyle cellStyle_RED;
	private XSSFCellStyle cellStyle_GRN;
	private XSSFCellStyle cellStyle_ORN;
	private XSSFCellStyle cellStyle_BLU;
	private XSSFCellStyle cellStyle;
	private ArrayList result1 = new ArrayList();
	private ArrayList result2 = new ArrayList();
	private String reportname;// 报表名称
	private InputStream inputStream;
	
	public OutputDataToExcel1(ArrayList result1, ArrayList result2,InputStream inputStream, String reportname) {
		// TODO Auto-generated constructor stub
		this.result1 = result1;
		this.result2 = result2;
		this.reportname = reportname;
		this.inputStream = inputStream;
		executeOperation();
	}

	public void executeOperation() {
		// TODO Auto-generated method stub
		try {
			book = new XSSFWorkbook(inputStream);
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		XSSFSheet sheet1 = book.getSheetAt(0);

		//////////// 设置分组显示上方/下方
		sheet1.setRowSumsBelow(false);
		sheet1.setRowSumsRight(false);
		sheet1.setRowSumsBelow(false);
		sheet1.setRowSumsRight(false);

		writeDataToSheet(sheet1, result1, true);
		
		if(book.getNumberOfSheets()>1) {
			XSSFSheet sheet2 = book.getSheetAt(1);
			if(sheet2!=null) {
				writeDataToSheet2(sheet2, result2, true);
			}
		}
	

		try {
			String fullFileName = FileUtil.getReportFileName(reportname);
			File file = new File(fullFileName);
			if (file.exists()) {
				file.delete();
				file = new File(fullFileName);
			}

			FileOutputStream fOut = new FileOutputStream(file);
			try {
				book.write(fOut);
				fOut.flush();
				fOut.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

			// 打开excel
			//openFile(fullFileName);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	private void openFile(String fullFileName) {
		// TODO Auto-generated method stub
		try {
			System.out.println("cmd /c call " +'"'+fullFileName+'"');
			Runtime.getRuntime().exec("cmd /c call " +'"'+fullFileName+'"');
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private void writeDataToSheet(XSSFSheet sheet, ArrayList result, boolean setGroup) {
		// TODO Auto-generated method stub
		cellStyle_RED = sheet.getWorkbook().createCellStyle();
		cellStyle_RED.setFillForegroundColor(IndexedColors.RED.getIndex());
		cellStyle_RED.setFillPattern(CellStyle.SOLID_FOREGROUND);
		
		cellStyle_GRN= sheet.getWorkbook().createCellStyle();
		cellStyle_GRN.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		cellStyle_GRN.setFillPattern(CellStyle.SOLID_FOREGROUND);
		
		cellStyle_ORN= sheet.getWorkbook().createCellStyle();
		cellStyle_ORN.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
		cellStyle_ORN.setFillPattern(CellStyle.SOLID_FOREGROUND);
		
		cellStyle_BLU= sheet.getWorkbook().createCellStyle();
		cellStyle_BLU.setFillForegroundColor(IndexedColors.SKY_BLUE.getIndex());
		cellStyle_BLU.setFillPattern(CellStyle.SOLID_FOREGROUND);
		
		cellStyle= sheet.getWorkbook().createCellStyle();	
		cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
		
	
		
		int row = 9;

		int len = result.size();

		for (int i = 0; i < len; i++) {
			String[] values = (String[]) result.get(i);
			for (int j = 0; j < values.length; j++) {
				setCell(sheet, values[j], row, j);
			}
			row++;
		}


	}
	private void writeDataToSheet2(XSSFSheet sheet, ArrayList result, boolean setGroup) {
		// TODO Auto-generated method stub
		Font font = book.createFont();
		font.setFontName("宋体");
		font.setFontHeightInPoints((short)14);
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);
		XSSFCellStyle cellStyle3 = sheet.getWorkbook().createCellStyle();	
		cellStyle3.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle3.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle3.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle3.setBorderTop(XSSFCellStyle.BORDER_THIN);
		cellStyle3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		cellStyle3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		cellStyle3.setFont(font);
		
		Font font2 = book.createFont();
		font2.setFontName("宋体");
		font2.setFontHeightInPoints((short)11);
		XSSFCellStyle cellStyle2 = sheet.getWorkbook().createCellStyle();	
		cellStyle2.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle2.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle2.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle2.setBorderTop(XSSFCellStyle.BORDER_THIN);
		cellStyle2.setFont(font2);
		
		int row = 2;

	    Map<String,Integer[]> map = (Map<String, Integer[]>) result.get(0);
		String title = (String) result.get(1) + "车型 生产要件检查数量汇总表  ";
		String date = "日期：" + (String) result.get(2);
		String state = (String) result.get(3);
		
		setStringCellAndStyle(sheet,title,0,0,null,Cell.CELL_TYPE_STRING);
		setStringCellAndStyle(sheet,date,0,7,null,Cell.CELL_TYPE_STRING);
        int index = 0;
		if(map.size()>0) {
			for(Map.Entry<String,Integer[]> entry: map.entrySet()) {
				String key = entry.getKey();
				Integer[] value = entry.getValue();
				Integer difference = value[0]-value[1];
				setStringCellAndStyle(sheet,key,row+index,0,null,Cell.CELL_TYPE_STRING);	
				setStringCellAndStyle(sheet,value[0].toString(),row+index,1,null,10);
				setStringCellAndStyle(sheet,state,row+index,2,null,Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet,value[1].toString(),row+index,3,null,10);
				setStringCellAndStyle(sheet,difference.toString(),row+index,4,null,10);
				setStringCellAndStyle(sheet,value[2].toString(),row+index,5,null,10);
				setStringCellAndStyle(sheet,"",row+index,6,null,Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet,"",row+index,7,null,Cell.CELL_TYPE_STRING);
				index++;
			}
		}		
	}

	// 对单元格赋值
		public static void setStringCellAndStyle(XSSFSheet sheet, String value, int rowIndex, int cellIndex,
				XSSFCellStyle Style, int celltype) {

			// 对于整型与字符型的区分 10为整型，11为double型

			XSSFRow row = sheet.getRow(rowIndex);
			if (row == null)
				row = sheet.createRow(rowIndex);
			XSSFCell cell = row.getCell(cellIndex);
			if (cell == null)
				cell = row.createCell(cellIndex);
			// cell.setCellType(celltype);
			if (value == null || value.isEmpty()) {
				cell.setCellType(Cell.CELL_TYPE_BLANK);
			} else {
				if (celltype == Cell.CELL_TYPE_STRING) {
					cell.setCellType(Cell.CELL_TYPE_STRING);
					cell.setCellValue(value);
				} else if (celltype == 10) {
					cell.setCellType(Cell.CELL_TYPE_NUMERIC);
					cell.setCellValue(Integer.parseInt(value));
				} else if (celltype == 11) {
					cell.setCellType(Cell.CELL_TYPE_NUMERIC);
					cell.setCellValue(Double.parseDouble(value));
				} else {
					cell.setCellType(Cell.CELL_TYPE_STRING);
					cell.setCellValue(value);
				}
			}
            if(Style!=null) {
            	cell.setCellStyle(Style);
            }
			

		}
		
	private void writeExcelHeader(XSSFSheet sheet, String[] title, int[] cloumnWidth) {
		// TODO Auto-generated method stub
		for (int i = 0; i < title.length; i++) {
			setCell(sheet, title[i], 0, i);
		}

		for (int i = 0; i < cloumnWidth.length; i++) {
			sheet.setColumnWidth(i, cloumnWidth[i]);
		}
	}
	
	// 对单元格赋值，强制使用文本格式
	protected void setCell(XSSFSheet sheet, String value, int rowIndex, int cellIndex) {
		XSSFRow row = sheet.getRow(rowIndex);
		if (row == null)
		{
			row = sheet.createRow(rowIndex);
		}
		XSSFCell cell = row.getCell(cellIndex);
		if (cell == null)
		{
			cell = row.createCell(cellIndex);
		}
		cell.setCellType(Cell.CELL_TYPE_STRING);
		// cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellValue(value);
		if(cellIndex != 0) {
		cell.setCellStyle(cellStyle);
		}
		if (cellIndex == 9) {
			if (value.equals("1")) {
				// 红色
				cell.setCellStyle(cellStyle_RED);
			}
			if (value.equals("2")) {
				// 绿色
				cell.setCellStyle(cellStyle_GRN);
			}
			if (value.equals("3")) {
				//  桔色
				cell.setCellStyle(cellStyle_ORN);
			}
			if (value.equals("5")) {
				//  蓝色
				cell.setCellStyle(cellStyle_BLU);
			}
		}
	}

}
