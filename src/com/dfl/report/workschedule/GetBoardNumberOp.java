package com.dfl.report.workschedule;

import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.swt.widgets.Shell;

import com.dfl.report.ExcelReader.WeldPointBoardInformation;
import com.dfl.report.ExcelReader.baseinfoExcelReader;
import com.dfl.report.util.FileUtil;
import com.dfl.report.util.GenerateReportInfo;
import com.dfl.report.util.NewOutputDataToExcel;
import com.dfl.report.util.ReportUtils;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentUser;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;

public class GetBoardNumberOp {

	private Shell shell;
	private TCSession session;
	private XSSFWorkbook book;
	private TCComponentBOMLine topbomline;
	private GenerateReportInfo info;
	private List<WeldPointBoardInformation> baseinfolist;// 基本信息表的数据

	public GetBoardNumberOp(TCSession session, XSSFWorkbook book, TCComponentBOMLine topbomline,
			GenerateReportInfo info, List<WeldPointBoardInformation> baseinfolist) throws TCException {
		// TODO Auto-generated constructor stub
		this.session = session;
		this.book = book;
		this.topbomline = topbomline;
		this.info = info;
		this.baseinfolist = baseinfolist;
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

		// 获取基本信息
//		String baseName = "222.基本信息";
//		baseinfolist = getBaseinfomation(topbomline.window().getTopBOMLine(), baseName);

		viewPanel.addInfomation("正在更新报表...\n", 30, 100);

		// 获取点焊sheet页
		List sheetlist = getSpotWeldingSheets();

		viewPanel.addInfomation("", 50, 100);

		// 根据焊点编号，获取对应的板组编号并写入
		writeBoradDataTOSheet(sheetlist);

		viewPanel.addInfomation("", 60, 100);

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

	private void writeBoradDataTOSheet(List sheetlist) {
		// TODO Auto-generated method stub
		if (sheetlist != null && sheetlist.size() > 0) {
			// 设置字体颜色
			Font font = book.createFont();
			font.setColor((short) 12);// 蓝色字体
			font.setFontName("MS PGothic");
			font.setFontHeightInPoints((short) 12);
			XSSFCellStyle style = book.createCellStyle();
			style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
			style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
			style.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
			// style.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
			style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style.setFont(font);

			XSSFCellStyle style2 = book.createCellStyle();
			style2.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
			// style2.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
			style2.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
			// style2.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
			style2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style2.setFont(font);

			XSSFCellStyle style3 = book.createCellStyle();
			style3.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
			// style3.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
			style3.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
			style3.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
			style3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style3.setFont(font);

			for (int i = 0; i < sheetlist.size(); i++) {
				String sheetname = (String) sheetlist.get(i);
				XSSFSheet sheet = book.getSheet(sheetname);
				for (int j = 0; j < 9; j++) {
					int index = 17 + 2 * j;
					XSSFRow row = sheet.getRow(index);
					if (row == null) {
						row = sheet.createRow(index);
					}
					XSSFCell cell;
					cell = row.getCell(116);
					String weldNo = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
					if (weldNo != null) {
						String[] board = getBoradCode(weldNo);
						setStringCellAndStyle(sheet, board[0], index, 103, style, Cell.CELL_TYPE_STRING);
						setStringCellAndStyle(sheet, board[1], index, 107, style2, Cell.CELL_TYPE_STRING);
						setStringCellAndStyle(sheet, board[2], index, 111, style3, Cell.CELL_TYPE_STRING);
					}
				}
			}
		}
	}

	/*
	 * 根据焊点号查找板组编号
	 */
	private String[] getBoradCode(String weldNo) {
		String[] borad = new String[3];
		if(baseinfolist!=null) {
			for (int i = 0; i < baseinfolist.size(); i++) {
				WeldPointBoardInformation wp = baseinfolist.get(i);
				if (weldNo.equals(wp.getWeldno())) {
					borad[0] = wp.getBoardnumber1();
					borad[1] = wp.getBoardnumber2();
					borad[2] = wp.getBoardnumber3();
					break;
				}
			}
		}
		
		return borad;
	}

	/*
	 * 获取点焊sheet
	 */
	private List getSpotWeldingSheets() {
		// TODO Auto-generated method stub
		int sheetnum = 0;
		List sheetList = new ArrayList();
		sheetnum = book.getNumberOfSheets();
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("点焊")) {
				sheetList.add(sheetname);
			}
		}
		return sheetList;
	}

	/*
	 * 获取基本信息表信息
	 */
	private List<WeldPointBoardInformation> getBaseinfomation(TCComponentBOMLine topbl, String procName) {
		List<WeldPointBoardInformation> baseinfolist = new ArrayList<WeldPointBoardInformation>();
		InputStream filein = null;
		try {
			filein = baseinfoExcelReader.getFileinbyreadExcel2(topbl.getItemRevision(), "IMAN_reference", procName);
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		baseinfolist = baseinfoExcelReader.readHDExcel(filein, "xlsx");

		return baseinfolist;
	}

	private static String convertCellValueToString(Cell cell, int type) {
		if (cell == null) {
			return null;
		}
		String returnValue = null;
		if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {

		} else {
			switch (type) {
			case Cell.CELL_TYPE_NUMERIC: // 数字
				Double doubleValue = cell.getNumericCellValue();
				// 格式化科学计数法，取一位整数
				DecimalFormat df = new DecimalFormat("0.0");
				returnValue = df.format(doubleValue);
				break;
			case Cell.CELL_TYPE_STRING: // 字符串
				cell.setCellType(Cell.CELL_TYPE_STRING);
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
		}
		return returnValue;
	}

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
}
