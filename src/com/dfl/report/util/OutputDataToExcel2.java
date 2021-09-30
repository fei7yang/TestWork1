package com.dfl.report.util;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLDecoder;
import java.util.Collections;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFBorderFormatting;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class OutputDataToExcel2 {
	private XSSFWorkbook book;
	private XSSFRow nowRow;
	private XSSFCellStyle cellStyle ;
	private XSSFCellStyle cellStyle1 ;
	private XSSFCell currentCell;
	private String[][] result1;
	private String reportname;// 报表名称
	private InputStream inputStream;
	private String vehicle;//车型
	
	public OutputDataToExcel2(String[][] result1, InputStream inputStream, String reportname, String vehicle) {
		// TODO Auto-generated constructor stub
		this.result1 = result1;
		this.reportname = reportname;
		this.inputStream = inputStream;
		this.vehicle=vehicle;
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
		XSSFSheet sheet1 = book.getSheet("清单");

		//////////// 设置分组显示上方/下方
		sheet1.setRowSumsBelow(false);
		sheet1.setRowSumsRight(false);
		sheet1.setRowSumsBelow(false);
		sheet1.setRowSumsRight(false);

		writeDataToSheet(sheet1, result1, true);

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
			System.out.println("cmd /c start " + "\"" + fullFileName + "\"");
			Runtime.getRuntime().exec("cmd /c " + fullFileName);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private void writeDataToSheet(XSSFSheet sheet, String[][] result, boolean setGroup) {
		// TODO Auto-generated method stub
		//写入车型的值
		
		setCell(sheet, vehicle, 1, 1);
		
		// 垂直居中
		cellStyle1=sheet.getWorkbook().createCellStyle();
		cellStyle1.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		cellStyle1.setAlignment(XSSFCellStyle.ALIGN_CENTER_SELECTION); // 创建一个居中格式
		cellStyle1.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle1.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle1.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle1.setBorderTop(XSSFCellStyle.BORDER_THIN);
		
		cellStyle= sheet.getWorkbook().createCellStyle();	
		cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
		int row = 4;

		int len = result.length;

		for (int i = 0; i < len; i++) {
			String[] values = result[i];
			for (int j = 0; j < values.length; j++) {
				setCell(sheet, values[j], row, j);
			}
			row++;
		}
		// 合并单元格

		// 从起始行开始向网下比较块的名称，将名称相同的块的单元格进行合并
		
		int currnetRow = 4;// 开始查找的起始行
		for (int p = 4; p < result1.length + 4; p++) {// 总行数
			currentCell = sheet.getRow(p).getCell(2);

			XSSFCell nextCell = null;
			String next = "";
			// 获取块名称
			String current = getStringCellValue(currentCell);
			if (p < result1.length + 5) {
				nowRow = sheet.getRow(p + 1);
				if (nowRow != null) {
					nextCell = nowRow.getCell(2);
					next = getStringCellValue(nextCell);
				} else {
					next = "";
				}
			} else {
				next = "";
			}
			if (current.equals(next)) {// 比对块的值是否相同
				currentCell.setCellValue(current);
				continue;
			} else {
				sheet.addMergedRegion(new CellRangeAddress(currnetRow, p, 2, 2));// 合并块前面的序号单元格				
				sheet.addMergedRegion(new CellRangeAddress(currnetRow, p, 3, 3));// 合并块单元格				
				currnetRow = p + 1;
			}
			
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
		if (row == null) {
			row = sheet.createRow(rowIndex);
		}
		XSSFCell cell = row.getCell(cellIndex);
		if (cell == null) {
			cell = row.createCell(cellIndex);
		}
		cell.setCellType(Cell.CELL_TYPE_STRING);
		// cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellValue(value);
		if(cellIndex!=0&&cellIndex!=1) {
			if(cellIndex==2||cellIndex==3||cellIndex==4) {
				cell.setCellStyle(cellStyle1);
			}else {
				cell.setCellStyle(cellStyle);
			}
			
		}
		
		
	}

	private String getStringCellValue(XSSFCell Cell) {
		// TODO Auto-generated method stub
		String strCell = "";
		if (Cell != null) {
			switch (Cell.getCellType()) {
			case XSSFCell.CELL_TYPE_STRING:
				strCell = Cell.getStringCellValue();
				break;
			case XSSFCell.CELL_TYPE_NUMERIC:
				strCell = String.valueOf(Cell.getNumericCellValue());
				break;
			case XSSFCell.CELL_TYPE_BOOLEAN:
				strCell = String.valueOf(Cell.getBooleanCellValue());
				break;
			case XSSFCell.CELL_TYPE_BLANK:
				strCell = "";
				break;
			default:
				strCell = "";
				break;
			}
			if (strCell.equals("") || strCell == null) {
				return "";
			}
			if (Cell == null) {
				return "";
			}
		}
		return strCell;
	}
}
