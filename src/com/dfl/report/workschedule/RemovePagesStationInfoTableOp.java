package com.dfl.report.workschedule;

import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.swt.widgets.Shell;

import com.dfl.report.util.FileUtil;
import com.dfl.report.util.GenerateReportInfo;
import com.dfl.report.util.NewOutputDataToExcel;
import com.dfl.report.util.ReportUtils;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentUser;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;

public class RemovePagesStationInfoTableOp {
	private Shell shell;
	private TCSession session;
	private XSSFWorkbook book;
	private TCComponentBOMLine topbomline;
	private List sheetlist;
	private GenerateReportInfo info;
	SimpleDateFormat df2 = new SimpleDateFormat("yyyy.MM");// 设置日期格式

	public RemovePagesStationInfoTableOp(TCSession session, XSSFWorkbook book, TCComponentBOMLine topbomline,
			List sheetlist, GenerateReportInfo info) throws TCException {
		// TODO Auto-generated constructor stub
		this.session = session;
		this.book = book;
		this.topbomline = topbomline;
		this.sheetlist = sheetlist;
		this.info = info;
		initUI();
	}

	private void initUI() throws TCException {
		// TODO Auto-generated method stub
		// 界面显示进度并输出执行步骤
		ReportViwePanel viewPanel = new ReportViwePanel("更新报表");
		viewPanel.setVisible(true);
		viewPanel.addInfomation("开始更新报表...\n", 20, 100);
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // PSW所在位置
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("有效页")) {
				sheetAtIndex = i;
				break;
			}
		}
		viewPanel.addInfomation("正在写入数据...\n", 40, 100);
		for (int i = 0; i < sheetlist.size(); i++) {
			String sheetname = (String) sheetlist.get(i);
			XSSFSheet sheet = book.getSheet(sheetname);
			int shindex = book.getSheetIndex(sheet);
			// 写更改记录信息和基本信息
			wirteDataToSheet(sheet);

			if (sheetAtIndex != -1) {
				// 获取减页的页码
				String pages = "";
				pages = getBaseinfomation(50, 107,shindex);
				// 修改有效页
				updateValidPage(sheetAtIndex, pages);
			}
		}
		String procName = Util.getProperty(info.getMeDocument(), "object_name");
		String filename = Util.formatString(procName);
		NewOutputDataToExcel.exportFile(book, filename);

		// 开启旁路
		{
			Util.callByPass(session, true);
		}
		viewPanel.addInfomation("", 80, 100);
		String fullFileName = FileUtil.getReportFileName(filename);
		System.out.println(fullFileName);
		TCComponentDataset ds = Util.createDataset(session, filename, fullFileName, "MSExcelX", "excel");
		List<TCComponentDataset> datasetList = new ArrayList<TCComponentDataset>();
		List<TCComponentItemRevision> revlist = new ArrayList<TCComponentItemRevision>();
		if (ds != null) {
			datasetList.add(ds);
		}
		revlist.add(topbomline.getItemRevision());
		try {
			TCComponentItem docunment = ReportUtils.afterGenerateReportAction(datasetList, revlist, info, procName, "",
					session);

		} catch (TCException e) {
			e.printStackTrace();
			return;
		}
		viewPanel.addInfomation("", 80, 100);
		// 关闭旁路
		{
			Util.callByPass(session, false);
		}

		viewPanel.addInfomation("报表更新完成，请在焊装工厂工位对象附件下查看报表...\n", 100, 100);

	}

	private void updateValidPage(int sheetAtIndex, String pages) {
		// TODO Auto-generated method stub
		// TODO Auto-generated method stub
		int page = 0;
		String edition = "";
		if (pages.length() > 0) {
			if(Util.isNumber(pages)){
				edition = "";
				page = Integer.parseInt(pages);
			}else {
				edition = pages.substring(pages.length() - 1);
				String str = pages.substring(0, pages.length() - 1);
				if (Util.isNumber(str)) {
					page = Integer.parseInt(str);
				}
			}			
		}
		// 设置字体颜色
		Font font = book.createFont();
		font.setColor((short) 12);// 蓝色字体
		font.setFontName("宋体");
		font.setFontHeightInPoints((short) 14);
		XSSFCellStyle style = book.createCellStyle();
		style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
		style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
		style.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
		style.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style.setFont(font);
		if (page != 0) {
			XSSFSheet sheet = book.getSheetAt(sheetAtIndex);
			int col = (page - 1) / 40;
			int firstRow = 7;
			int endRow = 47;
			int colindex = 6 + 35 * col;
			for (int i = firstRow; i < endRow; i++) {
				XSSFRow row = sheet.getRow(i);
				XSSFCell cell = row.getCell(colindex);
				String value = convertCellValueToString(cell);
				if (Util.isNumber(value)) {
					if (Double.parseDouble(value) == page) {
						if (edition.equals("A")) {
							setStringCellAndStyle(sheet, "", i, 15 + 35 * col, style, Cell.CELL_TYPE_STRING);
						} else if (edition.equals("B")) {
							setStringCellAndStyle(sheet, "", i, 19 + 35 * col, style, Cell.CELL_TYPE_STRING);
						} else if (edition.equals("C")) {
							setStringCellAndStyle(sheet, "", i, 23 + 35 * col, style, Cell.CELL_TYPE_STRING);
						} else if (edition.equals("D")) {
							setStringCellAndStyle(sheet, "", i, 27 + 35 * col, style, Cell.CELL_TYPE_STRING);
						} else if (edition.equals("E")) {
							setStringCellAndStyle(sheet, "", i, 31 + 35 * col, style, Cell.CELL_TYPE_STRING);
						} else if (edition.equals("F")) {
							setStringCellAndStyle(sheet, "", i, 35 + 35 * col, style, Cell.CELL_TYPE_STRING);
						} else {
							setStringCellAndStyle(sheet, "", i, 11 + 35 * col, style, Cell.CELL_TYPE_STRING);
						}
					}
				}

			}
		}
	}

	private void wirteDataToSheet(XSSFSheet sheet) throws TCException {
		// TODO Auto-generated method stub
		// 设置字体
		Font font3 = book.createFont();
		font3.setColor((short) 12);
		font3.setFontName("宋体");
		font3.setFontHeightInPoints((short) 12);
		// 创建一个样式
		XSSFCellStyle cellStyle3 = book.createCellStyle();
		cellStyle3.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle3.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle3.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle3.setBorderTop(XSSFCellStyle.BORDER_THIN);
		cellStyle3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
		cellStyle3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		cellStyle3.setFont(font3);
		TCComponentUser user = session.getUser();
		String username = user.getUserName();
		// 获取标记行是否有写
		String value = "";
		int row = 0;
		int col = 0;
		int sheetindex = book.getSheetIndex(sheet);
		value = getBaseinfomation(49, 3,sheetindex);
		if (value==null ||value.isEmpty()) {
			row = 49;
			col = 3;
		} else {
			value = getBaseinfomation(50, 3,sheetindex);
			if (value==null ||value.isEmpty()) {
				row = 50;
				col = 3;
			} else {
				value = getBaseinfomation(51, 3,sheetindex);
				if (value==null ||value.isEmpty()) {
					row = 51;
					col = 3;
				} else {
					value = getBaseinfomation(49, 35,sheetindex);
					if (value==null ||value.isEmpty()) {
						row = 49;
						col = 35;
					} else {
						value = getBaseinfomation(50, 35,sheetindex);
						if (value==null ||value.isEmpty()) {
							row = 50;
							col = 35;
						} else {
							value = getBaseinfomation(51, 35,sheetindex);
							if (value==null ||value.isEmpty()) {
								row = 51;
								col = 35;
							} else {

							}
						}
					}
				}
			}
		}
		setStringCellAndStyle(sheet, "取消", row, col, cellStyle3, Cell.CELL_TYPE_STRING);// 标记
		setStringCellAndStyle(sheet, "1", row, col+4, cellStyle3, Cell.CELL_TYPE_STRING);// 处数
		setStringCellAndStyle(sheet, username, row, col+20, cellStyle3, Cell.CELL_TYPE_STRING);// 签字
		setStringCellAndStyle(sheet, df2.format(new Date()), row, col+26, cellStyle3, Cell.CELL_TYPE_STRING);// 日期
	}

	/*
	 * 从生成的报表中获取 信息
	 */
	private String getBaseinfomation(int rowindex, int colindex, int sheetindex) {
		// TODO Auto-generated method stub
		XSSFCellStyle cellStyle2 = book.createCellStyle();
		Font font2 = book.createFont();
		font2.setColor(IndexedColors.BLUE.getIndex());
		font2.setBoldweight(Font.BOLDWEIGHT_BOLD); // 字体加粗
		font2.setFontHeightInPoints((short) 16);
		font2.setFontName("宋体");
		cellStyle2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
		cellStyle2.setAlignment(XSSFCellStyle.ALIGN_LEFT);
		cellStyle2.setFont(font2);
		String values = "";
		XSSFSheet sheet = book.getSheetAt(sheetindex);
		XSSFRow row;
		XSSFCell cell;
		row = sheet.getRow(rowindex);
		cell = row.getCell(colindex);
		values = convertCellValueToString(cell);// 页码	
		if(Util.isNumber(values)) {
			double doustr = Double.parseDouble(values);
			int intstr = (int)doustr;
			setStringCellAndStyle(sheet,Integer.toString(intstr),rowindex,colindex,cellStyle2,10);
		}
		return values;
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

		//cell.setCellStyle(Style);

	}

	private static String convertCellValueToString(Cell cell) {
		if (cell == null) {
			return null;
		}
		String returnValue = null;
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC: // 数字
			Double doubleValue = cell.getNumericCellValue();
			// 格式化科学计数法，取一位整数
			DecimalFormat df = new DecimalFormat("0");
			returnValue = df.format(doubleValue);
			break;
		case Cell.CELL_TYPE_STRING: // 字符串
			returnValue = cell.getStringCellValue();
			break;
		case Cell.CELL_TYPE_BOOLEAN: // 布尔
			Boolean booleanValue = cell.getBooleanCellValue();
			returnValue = booleanValue.toString();
			break;
		case Cell.CELL_TYPE_BLANK: // 空值
			break;
		case Cell.CELL_TYPE_FORMULA: // 公式
			returnValue = cell.getCellFormula();
			break;
		case Cell.CELL_TYPE_ERROR: // 故障
			break;
		default:
			break;
		}
		return returnValue;
	}

}
