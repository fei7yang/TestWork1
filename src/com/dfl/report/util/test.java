package com.dfl.report.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class test {

	public static void main(String[] args) {

//		Map<String, String> map = new TreeMap<String, String>();
//
//        map.put("P1 P2", "kfc");
//        map.put("P3 P5", "wnba");
//        map.put("L1 L2", "nba");
//        map.put("12", "cba");
//
//        Map<String, String> resultMap = sortMapByKey(map);    //按Key进行排序
//
//        for (Map.Entry<String, String> entry : resultMap.entrySet()) {
//            System.out.println(entry.getKey() + " " + entry.getValue());
//        }
//        
//        String str = "asd\\sdf";
//        String[] arr = str.split("\\\\");
//        System.out.println(arr[0] + " " + arr[1]);
//        System.out.println(str.replace("\\", "-"));
//        
//        ArrayList<String> list = new ArrayList<>();
//        list.add("1123");
//        System.out.println(list);
//        List<String> list2 = new ArrayList<>();
//        list2 = list;
//        System.out.println(list2);
//        
//        System.out.println(Util.isNumber("a120"));
//        System.out.println(Util.isNumber("0.10r"));
//        System.out.println(Util.isNumber("12.10a"));
//		String s = "76003";
//		String str = BigDecimal.valueOf(2000).subtract(BigDecimal.valueOf(Double.parseDouble("2000"))).toString();
////		String str = "20001";
//		System.out.println(str);
		
		
//		 StringBuilder str = new StringBuilder();
//	        str.append("12346");
//	        System.out.println(str.toString());
//	        str.delete(0,str.length());
//	        System.out.println(str.toString());
//
		File file = new File("C:\\Users\\pc\\Desktop\\test.xlsx");
		try {
			XSSFWorkbook book = new XSSFWorkbook(new FileInputStream(file));
			// 设置字体颜色
			Font font = book.createFont();
			font.setColor((short) 12);
			XSSFSheet sheet = book.getSheetAt(0);
			XSSFRow row = sheet.getRow(0);
			XSSFCell cell = row.getCell(1);
			XSSFCellStyle style = cell.getCellStyle();
			style.setFillForegroundColor(new XSSFColor(new java.awt.Color(255,199,206)));
			style.setFillPattern(CellStyle.SOLID_FOREGROUND);
			XSSFFont font2 = style.getFont();
			font.setFontName(font2.getFontName());
			font.setFontHeightInPoints(font2.getFontHeightInPoints());
			style.setFont(font);
			cell.setCellStyle(style);
			
			exportFile((XSSFWorkbook) book);

			System.out.println("finish！！");
			
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
//			XSSFCellStyle style = (XSSFCellStyle) book.createCellStyle();
//			style.setFont(font);
//			style.setBorderBottom(CellStyle.BORDER_HAIR); // 虚线边框
//			style.setBorderLeft(CellStyle.BORDER_HAIR); // 虚线边框
//			style.setBorderRight(CellStyle.BORDER_HAIR); // 虚线边框
//			style.setBorderTop(CellStyle.BORDER_HAIR); // 虚线边框
			
//			XSSFCellStyle style1 = (XSSFCellStyle) book.createCellStyle();
//			style1.setFont(font);
//			style1.setBorderBottom(CellStyle.BORDER_MEDIUM_DASH_DOT); // 虚线边框
//			style1.setBorderLeft(CellStyle.BORDER_MEDIUM_DASH_DOT); // 虚线边框
//			style1.setBorderRight(CellStyle.BORDER_MEDIUM_DASH_DOT); // 虚线边框
//			style1.setBorderTop(CellStyle.BORDER_MEDIUM_DASH_DOT); // 虚线边框
//
//			
//			XSSFCellStyle style2 = (XSSFCellStyle) book.createCellStyle();
//			style2.setFont(font);
//			style2.setBorderBottom(CellStyle.BORDER_DASH_DOT_DOT); // 虚线边框
//			style2.setBorderLeft(CellStyle.BORDER_DASH_DOT_DOT); // 虚线边框
//			style2.setBorderRight(CellStyle.BORDER_DASH_DOT_DOT); // 虚线边框
//			style2.setBorderTop(CellStyle.BORDER_DASH_DOT_DOT); // 虚线边框
//			
//			XSSFCellStyle style3 = (XSSFCellStyle) book.createCellStyle();
//			style3.setFont(font);
//			style3.setBorderBottom(CellStyle.BORDER_MEDIUM_DASHED); // 虚线边框
//			style3.setBorderLeft(CellStyle.BORDER_MEDIUM_DASHED); // 虚线边框
//			style3.setBorderRight(CellStyle.BORDER_MEDIUM_DASHED); // 虚线边框
//			style3.setBorderTop(CellStyle.BORDER_MEDIUM_DASHED); // 虚线边框
//			
//			XSSFCellStyle style4 = (XSSFCellStyle) book.createCellStyle();
//			style4.setFont(font);
//			style4.setBorderBottom(CellStyle.BORDER_SLANTED_DASH_DOT); // 虚线边框
//			style4.setBorderLeft(CellStyle.BORDER_SLANTED_DASH_DOT); // 虚线边框
//			style4.setBorderRight(CellStyle.BORDER_SLANTED_DASH_DOT); // 虚线边框
//			style4.setBorderTop(CellStyle.BORDER_SLANTED_DASH_DOT); // 虚线边框

			
//			XSSFCellStyle style5 = (XSSFCellStyle) book.createCellStyle();
//			style5.setFont(font);
//			style5.setBorderBottom(CellStyle.BORDER_DOTTED); // 虚线边框
//			style5.setBorderLeft(CellStyle.BORDER_DOTTED); // 虚线边框
//			style5.setBorderRight(CellStyle.BORDER_DOTTED); // 虚线边框
//			style5.setBorderTop(CellStyle.BORDER_DOTTED); // 虚线边框
			
//			XSSFCellStyle style6 = (XSSFCellStyle) book.createCellStyle();
//			style6.setFont(font);
//			style6.setBorderBottom(CellStyle.BORDER_DASH_DOT); // 虚线边框
//			style6.setBorderLeft(CellStyle.BORDER_DASH_DOT); // 虚线边框
//			style6.setBorderRight(CellStyle.BORDER_DASH_DOT); // 虚线边框
//			style6.setBorderTop(CellStyle.BORDER_DASH_DOT); // 虚线边框
//
//			
//			XSSFCellStyle style7 = (XSSFCellStyle) book.createCellStyle();
//			style7.setFont(font);
//			style7.setBorderBottom(CellStyle.BORDER_MEDIUM_DASH_DOT_DOT); // 虚线边框
//			style7.setBorderLeft(CellStyle.BORDER_MEDIUM_DASH_DOT_DOT); // 虚线边框
//			style7.setBorderRight(CellStyle.BORDER_MEDIUM_DASH_DOT_DOT); // 虚线边框
//			style7.setBorderTop(CellStyle.BORDER_MEDIUM_DASH_DOT_DOT); // 虚线边框
//						
			
//			Sheet sheet = book.getSheetAt(0);
//			setStringCellAndStyle2(sheet, "BORDER_HAIR", 1, 2, style, Cell.CELL_TYPE_STRING);
//			
//			setStringCellAndStyle2(sheet, "123", 3, 2, style1, Cell.CELL_TYPE_STRING);
//			
//			setStringCellAndStyle2(sheet, "123", 5, 2, style2, Cell.CELL_TYPE_STRING);
//			
//			setStringCellAndStyle2(sheet, "123", 7, 2, style3, Cell.CELL_TYPE_STRING);
//			
//			setStringCellAndStyle2(sheet, "123", 9, 2, style4, Cell.CELL_TYPE_STRING);
			
//			setStringCellAndStyle2(sheet, "BORDER_DOTTED", 3, 2, style5, Cell.CELL_TYPE_STRING);
//			
//			setStringCellAndStyle2(sheet, "123", 13, 2, style6, Cell.CELL_TYPE_STRING);
//			
//			setStringCellAndStyle2(sheet, "123", 15, 2, style7, Cell.CELL_TYPE_STRING);

//			exportFile((XSSFWorkbook) book);
//
//			System.out.println("finish！！");
//
//		} catch (FileNotFoundException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		} catch (IOException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
//		callVBSProgram("C:\\Users\\Boom\\Desktop\\tempreport","C:\\Users\\Boom\\Desktop\\555_HJZ_直材清单（焊技）_2020.08.1016时.xlsx","C:\\Users\\Boom\\Desktop\\UpdateSplitExcel.vbs","","ZC");
	}

	/**
	 * 使用 Map按key进行排序
	 * 
	 * @param map
	 * @return
	 */
	public static Map<String, String> sortMapByKey(Map<String, String> map) {
		if (map == null || map.isEmpty()) {
			return null;
		}

		Map<String, String> sortMap = new TreeMap<String, String>(new MapKeyComparator());

		sortMap.putAll(map);

		return sortMap;
	}

	// 255序列焊接条件设定表 电流电压
	private static double getCurrentandVoltage(double average) {
		// TODO Auto-generated method stub
		double fact = 0;
		double yushu = average % 0.5;
		if (yushu > 0) {
			fact = average + 0.5 - average % 0.5;
		} else {
			fact = average;
		}
		if (fact < 7) {
			fact = 7;
		}
		if (fact > 17) {
			fact = 17;
		}

		return fact;
	}

	public static void setStringCellAndStyle2(Sheet sheet, String value, int rowIndex, int cellIndex,
			XSSFCellStyle Style, int celltype) {

		// 对于整型与字符型的区分 10为整型，11为double型

		Row row = sheet.getRow(rowIndex);
		if (row == null)
			row = sheet.createRow(rowIndex);
		Cell cell = row.getCell(cellIndex);
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

		cell.setCellStyle(Style);

	}

	// 输出文件
	public static void exportFile(XSSFWorkbook book) {
		try {
			String fullFileName = "C:\\Users\\pc\\Desktop\\test.xlsx";
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
			// openFile(fullFileName);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	private static File[] callVBSProgram(String tempPath, String xlsFilePath, String vbsFilePath, String prefilename,
			String reporttype) {
		// TODO Auto-generated method stub
		String oupFilePath = tempPath + "output";
		File dirfile = new File(oupFilePath);
		if (!dirfile.exists()) {
			dirfile.mkdir();
		}
		final String command = "wscript  \"" + vbsFilePath + "\" \"" + xlsFilePath + "\" " + oupFilePath + " \""
				+ prefilename + "\" " + "\"" + "11" + "\" " + "\"" + reporttype + "\"";
		System.out.println(command);
		try {
			Process process = Runtime.getRuntime().exec(command);
			try {
				process.waitFor();
			} catch (InterruptedException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			System.out.println("finish");

			String oupFilePath2 = tempPath + "outputsuccess";

			File file = new File(oupFilePath2);
			if (file.exists()) {
				File[] files = file.listFiles();
//							for (int i = 0; i < files.length; i++) {
//								System.out.println("files:"+files[i].getPath());
//							}
				if (files != null) {
					return files;
				}
			} else {
				System.out.println("vbs拆分文件错误！");
			}

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return null;
	}
}
