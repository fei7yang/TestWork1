package com.dfl.report.handlers;

import java.awt.Color;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.imageio.ImageIO;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.common.usermodel.LineStyle;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.sl.usermodel.Shape;
import org.apache.poi.ss.usermodel.BorderFormatting;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Chart;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.charts.AxisCrosses;
import org.apache.poi.ss.usermodel.charts.AxisPosition;
import org.apache.poi.ss.usermodel.charts.ChartDataSource;
import org.apache.poi.ss.usermodel.charts.ChartLegend;
import org.apache.poi.ss.usermodel.charts.DataSources;
import org.apache.poi.ss.usermodel.charts.LegendPosition;
import org.apache.poi.ss.usermodel.charts.ScatterChartData;
import org.apache.poi.ss.usermodel.charts.ValueAxis;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFAnchor;
import org.apache.poi.xssf.usermodel.XSSFBorderFormatting;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingRule;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSheetConditionalFormatting;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFTextBox;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;

import com.dfl.report.util.NewOutputDataToExcel;
import com.dfl.report.util.OutputDataToExcel;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;

public class TestReport {

	public TestReport() {
		// TODO Auto-generated constructor stub
	}

	private static Map<String, Integer> map = new HashMap<String, Integer>();
	private static ArrayList list = new ArrayList();
	private static int index = 1;

	public static void main(String[] args) {
		// TODO Auto-generated method stub

		// ???????????????????????????????????????
//		ReportViwePanel viewPanel = new ReportViwePanel("????????????");
//		viewPanel.setVisible(true);
//		
//		
//		viewPanel.addInfomation("??????????????????...\n", 5, 100);
//		
//		
//		viewPanel.addInfomation("??????????????????...\n", 100, 100);
//        String str ="dadsass/ds\\dasd?--,";
//        String str2=str.replaceAll("[*|':;',\\\\[\\\\].<>/?~???@#???%??????&*????????????+|{}????????????????????????????????????]", "");
//        System.out.println(str2);
//		String path = "Z:\\?????????Eclipse??????\\imges";
//		 String fileStr3 = "new.png";
//	    String fileStr1 = "????????????1.png";
//	    String fileStr2 = "????????????2.png";
//	    String fileStr4 = "????????????3.png";
//	    String[] filestr = {"????????????1.png","????????????2.png","????????????3.png","????????????4.png","????????????5.png"};
//		try {
//			Util.mergeImage(fileStr1,fileStr2,fileStr4, fileStr3, path);
//		} catch (IOException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
//		System.out.println("?????????????????????"+path);

//	    try {
//			Util.batchmergeImage(filestr, fileStr3, path,200,200);
//			System.out.println("?????????????????????"+path);
//		} catch (IOException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}

//		FileOutputStream fileOut = null;
//		BufferedImage bufferImg = null;
//		BufferedImage bufferImg1 = null;
//		try {
//
//			// ????????????????????????????????????ByteArrayOutputStream??????????????????ByteArray
//			ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
//			ByteArrayOutputStream byteArrayOut1 = new ByteArrayOutputStream();
//			bufferImg = ImageIO.read(new File("Z:\\?????????Eclipse??????\\imges\\????????????1.png"));
//			bufferImg1 = ImageIO.read(new File("Z:\\?????????Eclipse??????\\imges\\????????????2.png"));
//			ImageIO.write(bufferImg, "png", byteArrayOut);
//			ImageIO.write(bufferImg1, "png", byteArrayOut1);
//
//			// ?????????????????????
//			XSSFWorkbook wb = new XSSFWorkbook();
//			XSSFSheet sheet1 = wb.createSheet("new sheet");
//			sheet1.setColumnWidth(0, 30 * 256);
////			sheet1.setDefaultColumnWidth(100*256);
////			sheet1.setDefaultRowHeight((short)(30*20));
//			// HSSFRow row = sheet1.createRow(2);
//           XSSFRow row = sheet1.createRow(0);
//           row.setHeight((short)2500);
//           Cell cell = row.createCell(0);
//           cell.setCellValue("???????????????");
//			// ??????????????????????????????????????????????????????????????????????????????????????????XSSFDrawing ?????????
//
//			XSSFDrawing patriarch = sheet1.createDrawingPatriarch();
//			XSSFClientAnchor anchor = new XSSFClientAnchor(0, 400*1000, 0,0, (short) 0, 0, (short) 1, 1);
//			XSSFClientAnchor anchor1 = new XSSFClientAnchor(0, 0, 100,100, (short) 0, 2, (short) 1, 3);
//			anchor1.setAnchorType(2);
//			anchor.setAnchorType(2);
//
		// ????????????
//			patriarch.createPicture(anchor, wb.addPicture(byteArrayOut.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
//			patriarch.createPicture(anchor1,
//					wb.addPicture(byteArrayOut1.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
//			XSSFSimpleShape shape1 = patriarch.createSimpleShape(anchor);  
//	        shape1.setShapeType(XSSFSimpleShape.EMU_PER_POINT);   
//	        shape1.setLineStyle(XSSFSimpleShape.POINT_DPI) ;  
		// ?????????
//			XSSFSheetConditionalFormatting f = sheet1.getSheetConditionalFormatting();  
//			XSSFConditionalFormattingRule r = f.createConditionalFormattingRule(ComparisonOperator.NOT_EQUAL, "\"NONE\"", null);  
//			XSSFBorderFormatting boderF = r.createBorderFormatting();  
//			boderF.setDiagonalBorderColor((short) 0);  
//			boderF.setBorderDiagonal(BorderFormatting.BORDER_THICK);  
//			boderF.setBottomBorderColor((short) 0);  
//			boderF.setBorderBottom(BorderFormatting.BORDER_THICK);  
//			          
//			XSSFConditionalFormattingRule[] rules = {r};  
//			CellRangeAddress[] regions = {new CellRangeAddress(0, 0, 0, 0)};  
//			f.addConditionalFormatting(regions, rules);
		// CellStyle style = wb.createCellStyle();
//			CellRangeAddress region1 = new CellRangeAddress(0, 2, (short) 0, (short) 0); //??????1???????????? ??????2???????????? ??????3???????????? ??????4???????????? 
//			sheet1.addMergedRegion(region1);
//	        XSSFCellStyle cellstyle = wb.createCellStyle();
//	        cellstyle.setVerticalAlignment(CellStyle.VERTICAL_TOP);
//	        cell.setCellStyle(cellstyle);
//			fileOut = new FileOutputStream("C:\\Users\\Administrator\\Desktop\\??????workbook.xlsx");
//			// ??????excel??????
//			wb.write(fileOut);
//			fileOut.close();
//			System.out.println("?????????????????????" + "C:\\Users\\Administrator\\Desktop\\??????workbook.xlsx");
//			NewOutputDataToExcel.openFile("C:\\Users\\Administrator\\Desktop\\??????workbook.xlsx");
//		} catch (IOException io) {
//			io.printStackTrace();
//			System.out.println("io erorr :  " + io.getMessage());
//
//		} finally {
//			if (fileOut != null) {
//				try {
//					fileOut.close();
//				} catch (IOException e) {
//					// TODO Auto-generated catch block
//					e.printStackTrace();
//				}
//			}
//		}
		File file = new File("C:\\Users\\Administrator\\Desktop\\????????????.xlsx");
	
			XSSFWorkbook book=null;
			try {
				book = new XSSFWorkbook(new FileInputStream(file));
			} catch (FileNotFoundException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			XSSFSheet sheet = book.getSheetAt(0);
//			XSSFRow row = sheet.getRow(1);
//			if(row == null) {
//				row = sheet.createRow(1);
//			}
//			XSSFCell cell = row.getCell(2);
//			if(cell==null) {
//				cell = row.createCell(2);
//			}
//			XSSFCell cell2 = row.getCell(4);
//			if(cell2==null) {
//				cell2 = row.createCell(4);
//			}
//			XSSFCell cell3 = row.getCell(6);
//			if(cell3==null) {
//				cell3 = row.createCell(6);
//			}
//			//??????????????????
//			Font font = book.createFont();
//			//font.setColor(Font.COLOR_RED);
//			font.setColor((short)12);
//			font.setFontName("??????");
//			CellStyle style = book.createCellStyle();
//			style.setFont(font);
//			//style.setBorderBottom(CellStyle.BORDER_DOUBLE); // ????????????
//			style.setBorderBottom(CellStyle.BORDER_HAIR); // ????????????
//			style.setBorderLeft(CellStyle.BORDER_HAIR); // ????????????
//			style.setBorderRight(CellStyle.BORDER_HAIR); // ????????????
//			style.setBorderTop(CellStyle.BORDER_HAIR); // ????????????

//			//style.setWrapText(true);
////			System.out.println("???????????????"+value);
////			System.out.println(value.getBytes());
//			cell.setCellType(Cell.CELL_TYPE_STRING);
//			cell.setCellValue("???");
//			cell.setCellStyle(style);
//			CellStyle style2 = book.createCellStyle();
//			style2.setBorderBottom(CellStyle.BORDER_DOUBLE); // ????????????
//			style2.setBorderLeft(CellStyle.BORDER_DOUBLE); // ????????????
//			style2.setBorderRight(CellStyle.BORDER_DOUBLE); // ????????????
//			style2.setBorderTop(CellStyle.BORDER_DOUBLE); // ????????????
//			cell2.setCellValue("???");
//			cell2.setCellStyle(style2);
//			
//			CellStyle style3 = book.createCellStyle();
//			style3.setBorderBottom(CellStyle.BORDER_THICK); // ????????????
//			style3.setBorderLeft(CellStyle.BORDER_THICK); // ????????????
//			style3.setBorderRight(CellStyle.BORDER_THICK); // ????????????
//			style3.setBorderTop(CellStyle.BORDER_THICK); // ????????????
//			cell3.setCellValue("???");
//			cell3.setCellStyle(style3);
//			
//			String value = cell.getStringCellValue();
		// cell.setCellType(Cell.CELL_TYPE_BLANK);
//		File newfile = new File("C:\\Users\\Administrator\\Desktop\\????????????.xlsx");
//		FileInputStream filein;
//		XSSFWorkbook book = null;
//		Workbook workbook = null;
//		try {
//			filein = new FileInputStream(newfile);
//			try {
//				book = new XSSFWorkbook(filein);
//				XSSFSheet sheet = book.getSheetAt(0);
//					List<Map<String, PictureData>> sheetList = new ArrayList<Map<String, PictureData>>(); 
//					Map<String, PictureData> sheetIndexPicMap = getSheetPictrues07(0, sheet, book);
//					sheetList.add(sheetIndexPicMap);
////					for(Map.Entry<String, PictureData> entry:sheetIndexPicMap.entrySet()) {
////						PictureData pic = entry.getValue();
////						
////					}
//					printImg(sheetList);

//					XSSFRow row = sheet.getRow(0);
//					XSSFCell cell0 = row.getCell(0);
//					if(cell0 == null) {
//						cell0 = row.createCell(0);
//					}
//			        cell0.setCellValue("??????            ??????") ;  
//			        //??????(???????????????????????????)  ???A1????????????cell?????????  ?????????????????????????????? 
//			        XSSFDrawing patriarch = sheet.createDrawingPatriarch();  
//			        XSSFClientAnchor a = new XSSFClientAnchor(0, 0, 1023, 255, (short)0, 0, (short)0, 0);  
//			        XSSFSimpleShape shape1 = patriarch.createSimpleShape(a);  
//			        shape1.setShapeType(ShapeTypes.LINE);   
//			        shape1.setLineStyle(ShapeTypes.ACCENT_BORDER_CALLOUT_1) ;
		// ShapeTypes st = new ShapeTypes();
//						Font font = book.createFont();
//						font.setColor((short)12);
//						CellStyle style = book.createCellStyle();
//						style.setFont(font);	
//						style.setFillForegroundColor(IndexedColors.SKY_BLUE.getIndex());
//						style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//						XSSFRow row = sheet.getRow(28);
//						if(row==null) {
//							row = sheet.createRow(28);
//						}
//						XSSFCell cell = row.getCell(14);
//						if(cell==null) {
//							cell=row.createCell(0);
//						}
//						cell.setCellStyle(style);														

//					int sheetIndex = book.getSheetIndex(sheet);
//					book.setPrintArea(sheetIndex, 0, 115, 0, 51);
//					
//					PrintSetup printSetup = sheet.getPrintSetup();
//					printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
//					printSetup.setScale((short)65);//????????????????????????100????????????
//					printSetup.setLandscape(true); // ???????????????true????????????false?????????(??????)   

		// sheet.removeMergedRegion(getMergedRegionIndex(sheet,7,4));

//				final int NUM_OF_ROWS = 3;
//				final int NUM_OF_COLUMNS = 10;
//
//				// Create a row and put some cells in it. Rows are 0 based.
//				Row row;
//				Cell cell;
//				for (int rowIndex = 0; rowIndex < NUM_OF_ROWS; rowIndex++) {
//					row = sheet.createRow((short) rowIndex);
//					for (int colIndex = 0; colIndex < NUM_OF_COLUMNS; colIndex++) {
//						cell = row.createCell((short) colIndex);
//						cell.setCellValue(colIndex * (rowIndex + 1));
//					}
//				}
//
//				Drawing drawing = sheet.createDrawingPatriarch();
//				ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 2, 5, 10, 15);// ????????????
//
//				Chart chart = drawing.createChart(anchor);// ????????????
//				ChartLegend legend = chart.getOrCreateLegend();// ?????????????????????
//				legend.setPosition(LegendPosition.TOP_RIGHT);// ??????????????????
//
//				ScatterChartData data = chart.getChartDataFactory().createScatterChartData();// ??????????????????
//
//				ValueAxis bottomAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.BOTTOM);// ??????
//				ValueAxis leftAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);// ??????
//				leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);// ??????????????????
//
//				ChartDataSource<Number> xs = DataSources.fromNumericCellRange(sheet,
//						new CellRangeAddress(0, 0, 0, NUM_OF_COLUMNS - 1));
//				ChartDataSource<Number> ys1 = DataSources.fromNumericCellRange(sheet,
//						new CellRangeAddress(1, 1, 0, NUM_OF_COLUMNS - 1));
//				ChartDataSource<Number> ys2 = DataSources.fromNumericCellRange(sheet,
//						new CellRangeAddress(2, 2, 0, NUM_OF_COLUMNS - 1));
//
//				data.addSerie(xs, ys1);
//				data.addSerie(xs, ys2);
//
//				chart.plot(data, bottomAxis, leftAxis);

		// ????????????????????????????????????
		// XSSFDrawing patriarch = sheet.createDrawingPatriarch();
		/**
		 * dx1?????????????????????x???????????????????????????255???????????????????????????A1???????????????????????????
		 * dy1?????????????????????y???????????????????????????125???????????????????????????A1???????????????????????????
		 * dx2?????????????????????x???????????????????????????1023???????????????????????????C3???????????????????????????
		 * dy2?????????????????????y???????????????????????????150???????????????????????????C3??????????????????????????? colFrom?????????????????????????????????0???????????????
		 * rowFrom?????????????????????????????????0???????????????????????????col1=0,row1=0???????????????????????????A1??? colTo?????????????????????????????????0???????????????
		 * rowTo?????????????????????????????????0???????????????????????????col2=2,row2=2???????????????????????????C3???
		 */
		// default
//				int dx1 = 0, dy1 = 0, dx2 = 1023, dy2 = 255;
//				int colFrom = 0, rowFrom = 0, colTo = 5, rowTo = 5;

		// XSSFClientAnchor bigValueAnchorShape = new XSSFClientAnchor(dx1, dy1, dx2,
		// dy2, (short)(colFrom), rowFrom, (short)(colTo), rowTo);
//				XSSFClientAnchor bigValueAnchorTextBox = new XSSFClientAnchor(dx1, dy1, dx2, dy2, (short)(colFrom+1), rowFrom+1, (short)(colTo-1), rowTo-1);
//				XSSFTextBox bigValueTextbox = patriarch.createTextbox(bigValueAnchorTextBox);
//				XSSFRichTextString str = new XSSFRichTextString("??????");
//				bigValueTextbox.setText(str );
//				bigValueTextbox.setFillColor(180, 205, 160);

//				XSSFClientAnchor line_anchor = new XSSFClientAnchor((short) 0, 0, dx1, dy1, (short) 1, 1, dx2, dy2);
//				XSSFSimpleShape line_shape = patriarch.createSimpleShape(line_anchor);
//				line_shape.setLineStyle(XSSFShape.EMU_PER_POINT);
//				line_shape.setNoFill(false);
//				// line_shape.setShapeType(XSSFSimpleShape.PIXEL_DPI);
//				line_shape.setLineWidth(1/12700);
//			} catch (IOException e) {
//				// TODO Auto-generated catch block
//				e.printStackTrace();
//			}
//		} catch (FileNotFoundException e2) {
//			// TODO Auto-generated catch block
//			e2.printStackTrace();
//		}
//		if (newfile.exists()) {
//			newfile.delete();
//			newfile = new File("C:\\Users\\Administrator\\Desktop\\????????????.xlsx");
//		}

		// ?????????????????????

//					File file = new File("???D:\\DFL_Template_DefectsCheckList.xlsx");
//					FileInputStream filein = new FileInputStream(file);
//					XSSFWorkbook wb = new XSSFWorkbook(filein);
//					XSSFSheet sheet = wb.getSheetAt(0);
//					CellRangeAddress region1 = new CellRangeAddress(1, 10, (short)6, (short)15); //??????1???????????? ??????2???????????? ??????3???????????? ??????4???????????? 
//					sheet.addMergedRegion(region1);
		FileOutputStream fOut = null;
		try {
			fOut = new FileOutputStream("C:\\Users\\Administrator\\Desktop\\????????????.xlsx");
		} catch (FileNotFoundException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		try {
			book.write(fOut);
			fOut.flush();
			fOut.close();

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		try {
			Runtime.getRuntime().exec("cmd /c C:\\Users\\Administrator\\Desktop\\????????????.xlsx");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		// java String?????????????????????????????????
//		String textContent = "??? abc test???";
//		
//		String str = getRemoveSpaces(textContent);
//		
//		System.out.println(str);
		/*
		 * ArrayList list = new ArrayList(); String[] str = new String[3]; //str[0] =
		 * "2" ; list.add(str); String[] str1 = new String[3]; str1[0] = "3" ;
		 * list.add(str1); String[] str2 = new String[3]; str2[0] = "1" ;
		 * list.add(str2); Comparator comparator = getComParatorBySerialID();
		 * Collections.sort(list, comparator); for(int i=0;i<list.size();i++) { String[]
		 * value = (String[]) list.get(i); System.out.println(value[0]); }
		 */

		Map<String, String> aftermap = new HashMap<String, String>();
		Map<String, String[]> premap = new HashMap<String, String[]>();

		aftermap.put("A", "B");
		aftermap.put("B", "C");
		aftermap.put("C", "D");
		aftermap.put("E", "F");
		aftermap.put("G", "F");
		aftermap.put("H", "F");
		aftermap.put("F", "I");
		aftermap.put("I", "J");
		aftermap.put("J", "D");
		aftermap.put("L", "K");
		aftermap.put("K", "D");

		premap.put("A", null);
		premap.put("B", new String[] { "A" });
		premap.put("C", new String[] { "B" });
		premap.put("D", new String[] { "C", "J", "K" });
		premap.put("E", null);
		premap.put("F", new String[] { "E", "G", "H" });
		premap.put("G", null);
		premap.put("H", null);
		premap.put("I", new String[] { "F" });
		premap.put("J", new String[] { "I" });
		premap.put("K", new String[] { "L" });
		premap.put("L", null);

		map.put("A", index);
		getNumber("A", premap, aftermap);
		System.out.println(map);
	}

	private static void getNumber(String first, Map<String, String[]> premap, Map<String, String> aftermap) {

		String after = aftermap.get(first);
		if (after == null) {
			return;
		}
		String[] pre = premap.get(after);

		if (pre != null && pre.length > 1) {
			if (!list.contains(after)) {
				list.add(after);
				getnoPre(first, pre, premap, aftermap);
			} else {
				if (first.equals(pre[pre.length - 1])) {
					if (!map.containsKey(after)) {
						index++;
						map.put(after, index);
						getNumber(after, premap, aftermap);
					}
				}
			}
		} else {
			if (!map.containsKey(after)) {
				index++;
				map.put(after, index);
				getNumber(after, premap, aftermap);
			}

		}

	}

	private static void getnoPre(String first, String[] pre, Map<String, String[]> premap,
			Map<String, String> aftermap) {

		for (int i = 0; i < pre.length; i++) {
			if (!pre[i].equals(first)) {
				System.out.println(pre[i]);
				String[] val = premap.get(pre[i]);
				if (val == null) {
					if (!map.containsKey(pre[i])) {
						index++;
						map.put(pre[i], index);
						getNumber(pre[i], premap, aftermap);
					}
				} else {
					String[] pre1 = premap.get(pre[i]);
					getnoPre(pre[i], pre1, premap, aftermap);
				}
			}
		}

	}

	private static java.util.List moveUp(ArrayList list, int startindex, int endindex) {
		java.util.List temp = new ArrayList();
		temp = list.subList(startindex, endindex + 1);
		System.out.println(temp);
		return temp;
	}

	private static Comparator getComParatorBySerialID() {
		// TODO Auto-generated method stub
		Comparator comparator = new Comparator() {

			public int compare(Object obj, Object obj1) {
				Object[] comp1 = (Object[]) obj;
				Object[] comp2 = (Object[]) obj1;
				int d1 = 0;
				int d2 = 0;
				if (comp1[0] != null && !comp1[0].toString().isEmpty()) {
					d1 = Integer.parseInt(comp1[0].toString());
				}
				if (comp2[0] != null && !comp2[0].toString().isEmpty()) {
					d2 = Integer.parseInt(comp2[0].toString());
				}
				if (d2 > d1) {
					return -1;
				}

				return 1;
			}
		};

		return comparator;
	}

	// String?????????????????????????????????
	private static String getRemoveSpaces(String str) {
		String afterStr = str;
		afterStr = afterStr.trim();

		while (afterStr.startsWith("???")) {// ?????????????????????????????????
			afterStr = afterStr.substring(1, afterStr.length()).trim();
		}
		while (afterStr.endsWith("???")) {
			afterStr = afterStr.substring(0, afterStr.length() - 1).trim();
		}
		return afterStr;
	}

	/**
	 * ???????????? Region
	 * 
	 * @param sheet
	 * @param row
	 * @param column
	 * @return
	 */
	public static int getMergedRegionIndex(XSSFSheet sheet, int row, int column) {
		int sheetMergeCount = sheet.getNumMergedRegions();

		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress ca = sheet.getMergedRegion(i);
			int firstColumn = ca.getFirstColumn();
			int lastColumn = ca.getLastColumn();
			int firstRow = ca.getFirstRow();
			int lastRow = ca.getLastRow();

			if (row >= firstRow && row <= lastRow) {
				if (column >= firstColumn && column <= lastColumn) {
					return i;
				}
			}
		}

		return 0;
	}

	/**
	 * ??????Excel2007??????
	 * 
	 * @param sheetNum ??????sheet??????
	 * @param sheet    ??????sheet??????
	 * @param workbook ???????????????
	 * @return Map key:????????????????????????0_1_1???String???value:?????????PictureData
	 */
	public static Map<String, PictureData> getSheetPictrues07(int sheetNum, XSSFSheet sheet, XSSFWorkbook workbook) {
		Map<String, PictureData> sheetIndexPicMap = new HashMap<String, PictureData>();

		for (POIXMLDocumentPart dr : sheet.getRelations()) {
			if (dr instanceof XSSFDrawing) {
				XSSFDrawing drawing = (XSSFDrawing) dr;
				List<XSSFShape> shapes = drawing.getShapes();
				for (XSSFShape shape : shapes) {
					XSSFPicture pic = (XSSFPicture) shape;
					XSSFClientAnchor anchor = pic.getPreferredSize();
					CTMarker ctMarker = anchor.getFrom();
					String picIndex = String.valueOf(sheetNum) + "_" + ctMarker.getRow() + "_" + ctMarker.getCol();
					sheetIndexPicMap.put(picIndex, pic.getPictureData());
				}
			}
		}
		return sheetIndexPicMap;
	}

	public static void printImg(List<Map<String, PictureData>> sheetList) throws IOException {

		for (Map<String, PictureData> map : sheetList) {
			Object key[] = map.keySet().toArray();
			for (int i = 0; i < map.size(); i++) {
				// ???????????????
				PictureData pic = map.get(key[i]);
				// ??????????????????
				String picName = key[i].toString();
				// ??????????????????
				String ext = pic.suggestFileExtension();

				byte[] data = pic.getData();

				FileOutputStream out = new FileOutputStream(
						"C:\\Users\\Administrator\\Desktop\\pic" + picName + "." + ext);
				out.write(data);
				out.close();
			}
		}

	}

	private static String getString(String str) {
		String value = "";

		return value;
	}
}
