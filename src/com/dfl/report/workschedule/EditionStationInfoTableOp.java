package com.dfl.report.workschedule;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
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
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentGroup;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentUser;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCProperty;
import com.teamcenter.rac.kernel.TCSession;

public class EditionStationInfoTableOp {
	private Shell shell;
	private TCSession session;
	private XSSFWorkbook book;
	private TCComponentBOMLine topbomline;
	private GenerateReportInfo info;
	SimpleDateFormat df2 = new SimpleDateFormat("yyyy.MM");// 设置日期格式
	public String Edition;

	public EditionStationInfoTableOp(TCSession session, XSSFWorkbook book, TCComponentBOMLine topbomline,
			GenerateReportInfo info, String edition) throws TCException {
		// TODO Auto-generated constructor stub
		this.session = session;
		this.book = book;
		this.topbomline = topbomline;
		this.info = info;
		this.Edition = edition;
		initUI();
	}

	private void initUI() throws TCException {
		// TODO Auto-generated method stub

		// 界面显示进度并输出执行步骤
		ReportViwePanel viewPanel = new ReportViwePanel("更新报表");
		viewPanel.setVisible(true);
		viewPanel.addInfomation("开始更新报表...\n", 20, 100);

		TCComponentUser user = session.getUser();
		String username = user.getUserName();
		String date = df2.format(new Date());
		TCComponentGroup group = session.getGroup();
		// 科室
		String groupname = group.getLocalizedFullName();
		// 发行科
		String department = "";
		if (groupname != null && (groupname.contains("同期工程科") || groupname.contains("simultaneous Engineering Section"))) {
			department = "H30";
		} else if (groupname != null && (groupname.contains("焊装技术科") || groupname.contains("Body Assembly Engineering Section"))) {
			department = "VE2";
		} else {
			department = "VE2";
		}

		viewPanel.addInfomation("正在更新报表...\n", 30, 100);
		// 有效页重新输出
		dealValidPage();

		// 更改标记清除，“编制”、“日期”重新输出
		dealClearAndwriteDateToSheet(username, date,department,Edition);

		viewPanel.addInfomation("", 50, 100);

		// 页码重排
		dealPageRearrangement();

		viewPanel.addInfomation("", 60, 100);

		// 开启旁路
		{
			Util.callByPass(session, true);
		}
		int shs = book.getNumberOfSheets();
		setPropertyValue(topbomline.getItemRevision(), "b8_OpSheetNumber", Integer.toString(shs));

		String procName = Util.getProperty(info.getMeDocument(), "object_name");
		String filename = Util.formatString(procName);
		NewOutputDataToExcel.exportFile(book, filename);

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
			// saveFileToFolder(docunment, topfoldername, childrenFoldername);

		} catch (TCException e) {
			e.printStackTrace();
			// EclipseUtils.info(Messages.FixtureOperation_15 + e.getMessage());
			return;
		}
		viewPanel.addInfomation("", 80, 100);
		// 关闭旁路
		{
			Util.callByPass(session, false);
		}

		viewPanel.addInfomation("报表更新完成，请在焊装工厂工位对象附件下查看报表...\n", 100, 100);
	}

	/*
	 * 设置属性值
	 */
	private void setPropertyValue(TCComponent tcc, String property, String value) throws TCException {
		TCProperty p = tcc.getTCProperty(property);
		if (p != null) {
			p.setStringValue(value);
		}
	}

	/*
	 * 页码重排
	 */
	private void dealPageRearrangement() {
		// TODO Auto-generated method stub
		int sheetnum = book.getNumberOfSheets();
		// 定义比较名称,初始值为第一个sheet名称，如果名称相同，则需要在名称后面增加1,2......
		String tempname = "";
		String sheetAllname;
		int num = 1;
		Pattern p = Pattern.compile("[0-9a-fA-F]"); // 非数字字母
		Pattern p2 = Pattern.compile("[0-9]"); // 非数字字母
		// 先按照顺序命令，避免后续按照规则命名出现名称重复情况
		for (int i = 0; i < sheetnum; i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			String sheetname = sheet.getSheetName();
			sheetAllname = sheetname + (i + 1);
			book.setSheetName(i, sheetAllname);
		}
		for (int i = 0; i < sheetnum; i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			// 截取前三位，去掉字母
			String sheetname = "";
			String oldsheetname = sheet.getSheetName();
			if (oldsheetname.length() > 2) {
				Matcher m = p.matcher(oldsheetname.substring(0, 3));
				Matcher m2 = p2.matcher(oldsheetname.substring(3));
				sheetname = m.replaceAll("").trim() + m2.replaceAll("").trim();
			} else {
				sheetname = oldsheetname;
			}
			// 第一个先不比较，从第二个开始和第一个比较
			if (i == 0) {
				tempname = sheetname;
				sheetAllname = String.format("%02d", i + 1) + sheetname;
				book.setSheetName(i, sheetAllname);

			} else {
				if (sheetname.contains(tempname)) {
					// 如果num为1，则说明sheet同名的第一个需要重命名增加数字后缀
					if (num == 1) {
						sheetAllname = String.format("%02d", i) + tempname + num;
						book.setSheetName(i - 1, sheetAllname);
					}
					sheetAllname = String.format("%02d", i + 1) + tempname + (num + 1);
					book.setSheetName(i, sheetAllname);
					num++;
				} else {
					sheetAllname = String.format("%02d", i + 1) + sheetname;
					tempname = sheetname;
					book.setSheetName(i, sheetAllname);

					num = 1;
				}
			}
			// 设置打印区域
			book.removePrintArea(i);
		}

		for (int i = 0; i < sheetnum; i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			book.setPrintArea(i, 0, 114, 0, 51);
			PrintSetup printSetup = sheet.getPrintSetup();
			printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
			printSetup.setScale((short) 70);// 自定义缩放，此处100为无缩放
			printSetup.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)
		}
	}

	/*
	 * 更改标记清除，“编制”、“日期”重新输出
	 */
	private void dealClearAndwriteDateToSheet(String username, String date, String department, String edition) {
		// TODO Auto-generated method stub
		// 设置字体
//		Font font = book.createFont();
//		font.setColor((short) 12);
//		font.setFontName("宋体");
//		font.setFontHeightInPoints((short) 16);
		// 创建一个样式
		XSSFCellStyle cellStyle1 = null;
//		cellStyle1.setBorderBottom(XSSFCellStyle.BORDER_THIN);
//		cellStyle1.setBorderLeft(XSSFCellStyle.BORDER_THIN);
//		cellStyle1.setBorderRight(XSSFCellStyle.BORDER_THIN);
//		cellStyle1.setBorderTop(XSSFCellStyle.BORDER_THIN);
//		cellStyle1.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
//		cellStyle1.setAlignment(XSSFCellStyle.ALIGN_CENTER);
//		cellStyle1.setFont(font);

		XSSFCellStyle cellStyle2 = null;
//		Font font2 = book.createFont();
//		font2.setColor(IndexedColors.BLUE.getIndex());
//		font2.setBoldweight(Font.BOLDWEIGHT_BOLD); // 字体加粗
//		font2.setFontHeightInPoints((short) 16);
//		font2.setFontName("宋体");
//		cellStyle2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
//		cellStyle2.setAlignment(XSSFCellStyle.ALIGN_LEFT);
//		cellStyle2.setFont(font2);

		// 设置字体
//		Font font3 = book.createFont();
//		font3.setColor((short) 12);
//		font3.setFontName("宋体");
//		font3.setFontHeightInPoints((short) 12);
		// 创建一个样式
		XSSFCellStyle cellStyle3 = null;
//		cellStyle3.setBorderBottom(XSSFCellStyle.BORDER_THIN);
//		cellStyle3.setBorderLeft(XSSFCellStyle.BORDER_THIN);
//		cellStyle3.setBorderRight(XSSFCellStyle.BORDER_THIN);
//		cellStyle3.setBorderTop(XSSFCellStyle.BORDER_THIN);
//		cellStyle3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
//		cellStyle3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
//		cellStyle3.setFont(font3);

		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		for (int i = 0; i < sheetnum; i++) {
			XSSFSheet sheet = book.getSheetAt(i);

			setStringCellAndStyle(sheet, department, 2, 0, cellStyle1, Cell.CELL_TYPE_STRING); // 发行科
			setStringCellAndStyle(sheet, username, 2, 6, cellStyle1, Cell.CELL_TYPE_STRING); // 编制
			setStringCellAndStyle(sheet, date, 2, 30, cellStyle1, Cell.CELL_TYPE_STRING);// 日期
			setStringCellAndStyle(sheet, Integer.toString(i + 1), 50, 107, cellStyle2, 10);// 当前页码
			setStringCellAndStyle(sheet, Integer.toString(sheetnum), 50, 112, cellStyle2, 10);// 总页码
//			setStringCellAndStyle(sheet, edition, 48, 108, cellStyle2, 10);// 批次

			setStringCellAndStyle(sheet, "", 49, 3, cellStyle3, Cell.CELL_TYPE_STRING);// 标记
			setStringCellAndStyle(sheet, "", 49, 7, cellStyle3, Cell.CELL_TYPE_STRING);// 处数
			setStringCellAndStyle(sheet, "", 49, 23, cellStyle3, Cell.CELL_TYPE_STRING);// 签字
			setStringCellAndStyle(sheet, "", 49, 29, cellStyle3, Cell.CELL_TYPE_STRING);// 日期

			setStringCellAndStyle(sheet, "", 50, 3, cellStyle3, Cell.CELL_TYPE_STRING);// 标记
			setStringCellAndStyle(sheet, "", 50, 7, cellStyle3, Cell.CELL_TYPE_STRING);// 处数
			setStringCellAndStyle(sheet, "", 50, 23, cellStyle3, Cell.CELL_TYPE_STRING);// 签字
			setStringCellAndStyle(sheet, "", 50, 29, cellStyle3, Cell.CELL_TYPE_STRING);// 日期

			setStringCellAndStyle(sheet, "", 51, 3, cellStyle3, Cell.CELL_TYPE_STRING);// 标记
			setStringCellAndStyle(sheet, "", 51, 7, cellStyle3, Cell.CELL_TYPE_STRING);// 处数
			setStringCellAndStyle(sheet, "", 51, 23, cellStyle3, Cell.CELL_TYPE_STRING);// 签字
			setStringCellAndStyle(sheet, "", 51, 29, cellStyle3, Cell.CELL_TYPE_STRING);// 日期

			setStringCellAndStyle(sheet, "", 49, 35, cellStyle3, Cell.CELL_TYPE_STRING);// 标记
			setStringCellAndStyle(sheet, "", 49, 39, cellStyle3, Cell.CELL_TYPE_STRING);// 处数
			setStringCellAndStyle(sheet, "", 49, 55, cellStyle3, Cell.CELL_TYPE_STRING);// 签字
			setStringCellAndStyle(sheet, "", 49, 61, cellStyle3, Cell.CELL_TYPE_STRING);// 日期

			setStringCellAndStyle(sheet, "", 50, 35, cellStyle3, Cell.CELL_TYPE_STRING);// 标记
			setStringCellAndStyle(sheet, "", 50, 39, cellStyle3, Cell.CELL_TYPE_STRING);// 处数
			setStringCellAndStyle(sheet, "", 50, 55, cellStyle3, Cell.CELL_TYPE_STRING);// 签字
			setStringCellAndStyle(sheet, "", 50, 61, cellStyle3, Cell.CELL_TYPE_STRING);// 日期

			setStringCellAndStyle(sheet, "", 51, 35, cellStyle3, Cell.CELL_TYPE_STRING);// 标记
			setStringCellAndStyle(sheet, "", 51, 39, cellStyle3, Cell.CELL_TYPE_STRING);// 处数
			setStringCellAndStyle(sheet, "", 51, 55, cellStyle3, Cell.CELL_TYPE_STRING);// 签字
			setStringCellAndStyle(sheet, "", 51, 61, cellStyle3, Cell.CELL_TYPE_STRING);// 日期

		}
	}

	/*
	 * 有效页重新输出
	 */
	private void dealValidPage() {
		// TODO Auto-generated method stub
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; //
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("有效页")) {
				sheetAtIndex = i;
				break;
			}
		}
		if (sheetAtIndex == -1) {
			return;
		}
		XSSFSheet sheet = book.getSheetAt(sheetAtIndex);
		// 设置字体颜色
//		Font font = book.createFont();
//		font.setColor((short) 12);// 蓝色字体
//		font.setFontName("宋体");
//		font.setFontHeightInPoints((short) 14);
		XSSFCellStyle style = null;
//		style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
//		style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
//		style.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
//		style.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
//		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
//		style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
//		style.setFont(font);
		// 先清空
		int page = 3;
		for (int i = 0; i < page; i++) {
			for (int j = 0; j < 40; j++) {
				for (int k = 0; k < 7; k++) {
					setStringCellAndStyle(sheet, "", 7 + j, 11 + 35 * i + k * 4, style, Cell.CELL_TYPE_STRING); //
				}
			}
		}
		// 再重新写
		page = (sheetnum - 1) / 40 + 1;
		for (int i = 0; i < page; i++) {
			if (i == page - 1) {
				for (int j = 0; j < sheetnum - 40 * i; j++) {
					setStringCellAndStyle(sheet, "●", 7 + j, 11 + 35 * i, style, Cell.CELL_TYPE_STRING); // 编制
				}
			} else {
				for (int j = 0; j < 40; j++) {
					setStringCellAndStyle(sheet, "●", 7 + j, 11 + 35 * i, style, Cell.CELL_TYPE_STRING); // 编制
				}
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
		if (Style != null) {
			cell.setCellStyle(Style);
		}

	}

}
