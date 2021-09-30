package com.dfl.report.workschedule;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
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
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.swt.widgets.Shell;

import com.dfl.report.util.CopySheetUtil;
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
import com.teamcenter.rac.kernel.TCProperty;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class AddPagesStationInfoTableOp {

	private Shell shell;
	private TCSession session;
	private XSSFWorkbook book;
	private TCComponentBOMLine topbomline;
	private String sheetname;
	private String newsheetname;
	private String sheetpages;
	private String model;
	private String modelname;
	private GenerateReportInfo info;
	private TCComponentDataset factdatawet;
	SimpleDateFormat df2 = new SimpleDateFormat("yyyy.MM");// 设置日期格式

	public AddPagesStationInfoTableOp(TCSession session, TCComponentDataset factdatawet, TCComponentBOMLine topbomline,
			String sheetname, String newsheetname, String sheetpages, String model, String modelname,
			GenerateReportInfo info) throws TCException, FileNotFoundException {
		// TODO Auto-generated constructor stub
		this.session = session;
		this.factdatawet = factdatawet;
		this.topbomline = topbomline;
		this.sheetname = sheetname;
		this.newsheetname = newsheetname;
		this.sheetpages = sheetpages;
		this.model = model;
		this.modelname = modelname;
		this.info = info;
		initUI();
	}

	private void initUI() throws TCException, FileNotFoundException {
		// TODO Auto-generated method stub

		// 界面显示进度并输出执行步骤
		ReportViwePanel viewPanel = new ReportViwePanel("更新报表");
		viewPanel.setVisible(true);
		viewPanel.addInfomation("开始更新报表...\n", 20, 100);

		// 将模板中选择的sheet复制到报表中
		TCComponentDataset dataset;
		if (model.equals("普通工位模板")) {
			dataset = FileUtil.getDatasetFile("DFL_Template_EngineeringWorkListStation");
			if (dataset == null) {
				viewPanel.addInfomation("错误：没有找到工程作业表普通工位模板，请先添加模板(名称为：DFL_Template_EngineeringWorkListStation)\\n",
						100, 100);
				return;
			}
		} else if (model.equals("VIN码打刻模板")) {
			dataset = FileUtil.getDatasetFile("DFL_Template_EngineeringWorkVINCarve");
			if (dataset == null) {
				viewPanel.addInfomation("错误：没有找到工程作业表VIN码打刻模板，请先添加模板(名称为：DFL_Template_EngineeringWorkVINCarve)\\n", 100,
						100);
				return;
			}
		} else {
			dataset = FileUtil.getDatasetFile("DFL_Template_AdjustmentLine");
			if (dataset == null) {
				viewPanel.addInfomation("错误：没有找到工程作业表调整线模板，请先添加模板(名称为：DFL_Template_AdjustmentLine)\\n", 100, 100);
				return;
			}
		}
		String tempPath = getTempPath();
		File dirfile = new File(tempPath);
		if (!dirfile.exists()) {
			dirfile.mkdir();
		}
		// 下载数据集文件
		File file = downloadFile((TCComponentDataset) dataset, tempPath);
		if (file == null) {
			System.out.println("下载数据集文件错误");
			viewPanel.addInfomation("下载数据集文件错误\n", 100, 100);
			return;
		}
		// 下载数据集文件
		File factfile = downloadFile(factdatawet, tempPath);
		if (factfile == null) {
			System.out.println("下载数据集文件错误");
			viewPanel.addInfomation("下载数据集文件错误\n", 100, 100);
			return;
		}
		// 下载VBS脚本
		File scriptFile = Util.getRCPPluginInsideFile("CopySheet.vbs");
		if (scriptFile == null || !scriptFile.exists()) {
			viewPanel.addInfomation("下载CopySheet.vbs脚本错误\n", 100, 100);
			// MessageBox.post("下载SplitExcel.vbs脚本错误", "工程作业表拆分", MessageBox.WARNING);
			return;
		}
		InputStream filein = new FileInputStream(file);
		XSSFWorkbook oldbook = NewOutputDataToExcel.creatXSSFWorkbook(filein);

		String modelsheetname = "";
		for (int i = 0; i < oldbook.getNumberOfSheets(); i++) {
			String shname = oldbook.getSheetName(i);
			if (shname.contains(modelname)) {
				modelsheetname = shname;
				break;
			}
		}

		String vbsFilePath = scriptFile.getAbsolutePath();
		// 调用本地vbs复制sheet
		File aferfile = callVBSProgram(tempPath, factfile.getAbsolutePath(), vbsFilePath, modelsheetname);
		viewPanel.addInfomation("", 40, 100);
		if (aferfile == null) {
			viewPanel.addInfomation("复制sheet失败\n", 100, 100);
			// MessageBox.post("下载SplitExcel.vbs脚本错误", "工程作业表拆分", MessageBox.WARNING);
			return;
		}
		InputStream afterfilein = new FileInputStream(aferfile);
		book = NewOutputDataToExcel.creatXSSFWorkbook(afterfilein);
		// f复制后的sheet在第一页
		XSSFSheet newsheet = book.getSheetAt(0);
		int index = book.getSheetIndex(sheetname);
		book.setSheetOrder(newsheet.getSheetName(), index);
		if (!newsheetname.contains(sheetpages)) {
			newsheetname = sheetpages + newsheetname;
		}
		try 
		{
			book.setSheetName(index, newsheetname);
		}
		catch(Exception e)
		{
			viewPanel.dispose();
			MessageBox.post("新增的sheet名称已存在，sheet名称不能重复！", "提示信息", MessageBox.INFORMATION);
			return ;
		}
		

		viewPanel.addInfomation("正在写入数据...\n", 60, 100);

		// 从生成的报表中获取 “编制”、“日期”、“版次”等信息信息
		String[] baseinfo = getBaseinfomation();

		// 写更改记录信息和基本信息
		wirteDataToSheet(newsheet, baseinfo);

		// 修改有效页
		updateValidPage();

		// 重新设置打印区域
		setPrintArea();

		// 获取作业内容
		int shs = book.getNumberOfSheets();
		String[] contents = new String[shs];
		for (int i = 0; i < shs; i++) {
			String sheetname = book.getSheetName(i);
			contents[i] = sheetname;
		}
		String procName = Util.getProperty(info.getMeDocument(), "object_name");
		String filename = Util.formatString(procName);
		NewOutputDataToExcel.exportFile(book, filename);
		// 开启旁路
		{
			Util.callByPass(session, true);
		}
		// 写入作业内容和焊点重要度
		TCProperty ppp = topbomline.getItemRevision().getTCProperty("b8_OperationContent");
		if (ppp != null) {
			ppp.setStringValueArray(contents);
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

		// 执行完后，把文件删除
		if (file.exists()) {
			file.delete();
		}
		if (factfile.exists()) {
			factfile.delete();
		}
		if (aferfile.exists()) {
			aferfile.delete();
		}

		viewPanel.addInfomation("报表更新完成，请在焊装工厂工位对象附件下查看报表...\n", 100, 100);
	}

	private File callVBSProgram(String tempPath, String absolutePath, String vbsFilePath, String sheetname) {
		// TODO Auto-generated method stub
		String oupFilePath = tempPath + "newreport";
		File dirfile = new File(oupFilePath);
		if (!dirfile.exists()) {
			dirfile.mkdir();
		}
		String newoupFilePath = oupFilePath + "\\newreport";
		final String command = "wscript  \"" + vbsFilePath + "\" \"" + absolutePath + "\" \"" + tempPath + " \" \""
				+ newoupFilePath + "\" \"" + sheetname + "\"";
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

			File file = new File(oupFilePath);
			if (file.exists()) {
				File[] files = file.listFiles();
				if (files != null && files.length > 0) {
					return files[0];
				} else {
					System.out.println("vbs复制sheet错误！");
				}

			} else {
				System.out.println("vbs复制sheet错误！");
			}

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return null;
	}

	/**
	 * 通过名称获取Word模板文件
	 * 
	 * @param name
	 * @return 数据集对象
	 */
	public File downloadFile(TCComponentDataset dataset, String difPath) {

		try {
			System.out.println(dataset.getType());
			if (dataset.getType().equals("MSExcelX")) {
				File files[] = dataset.getFiles("excel", difPath);
				if (files == null || files.length <= 0) {
					return null;
				}
				System.err.println(files[0].getPath());
				return files[0];
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;
	}

	public String getTempPath() {
		String path = "";
		String tmpPath = System.getProperty("java.io.tmpdir");
		// System.out.println("tmpPath:"+tmpPath);
		if (tmpPath.endsWith("\\")) {
			path = tmpPath + new Date().getTime();
		} else {
			path = tmpPath + "\\" + new Date().getTime();
		}
		path = path + "\\";
		System.out.println("tempPath=" + path);
		return path;
	}

	private void setPrintArea() {
		// TODO Auto-generated method stub
		int sheetnum = book.getNumberOfSheets();
		for (int i = 0; i < sheetnum; i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			book.setPrintArea(i, 0, 114, 0, 51);
			PrintSetup printSetup = sheet.getPrintSetup();
			printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
			printSetup.setScale((short) 70);// 自定义缩放，此处100为无缩放
			printSetup.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)
		}
	}

	private void updateValidPage() {
		// TODO Auto-generated method stub
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
		if (sheetAtIndex == -1) {
			return;
		}
		int page = 0;
		String edition = "";
		if (sheetpages.length() > 0) {
			edition = sheetpages.substring(sheetpages.length() - 1);
			String str = sheetpages.substring(0, sheetpages.length() - 1);
			if (Util.isNumber(str)) {
				page = Integer.parseInt(str);
			}
		}
		// 设置字体颜色
		XSSFCellStyle style = null;
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
							setStringCellAndStyle(sheet, "●", i, 15 + 35 * col, style, Cell.CELL_TYPE_STRING);
						} else if (edition.equals("B")) {
							setStringCellAndStyle(sheet, "●", i, 19 + 35 * col, style, Cell.CELL_TYPE_STRING);
						} else if (edition.equals("C")) {
							setStringCellAndStyle(sheet, "●", i, 23 + 35 * col, style, Cell.CELL_TYPE_STRING);
						} else if (edition.equals("D")) {
							setStringCellAndStyle(sheet, "●", i, 27 + 35 * col, style, Cell.CELL_TYPE_STRING);
						} else if (edition.equals("E")) {
							setStringCellAndStyle(sheet, "●", i, 31 + 35 * col, style, Cell.CELL_TYPE_STRING);
						} else if (edition.equals("F")) {
							setStringCellAndStyle(sheet, "●", i, 35 + 35 * col, style, Cell.CELL_TYPE_STRING);
						} else {

						}
					}
				}

			}
		}

	}

	/*
	 * 写更改记录信息和基本信息
	 */
	private void wirteDataToSheet(XSSFSheet newsheet, String[] baseinfo) throws TCException {
		// TODO Auto-generated method stub

		TCComponentUser user = session.getUser();
		String username = user.getUserName();

		// 页码第一位为0，不显示0，例如06A 显示为6A
		String page = "";
		if (sheetpages.substring(0, 1).equals("0")) {
			page = sheetpages.substring(1);
		} else {
			page = sheetpages;
		}

		setStringCellAndStyle(newsheet, baseinfo[0], 2, 6, null, Cell.CELL_TYPE_STRING); // 编制
		setStringCellAndStyle(newsheet, baseinfo[1], 2, 30, null, Cell.CELL_TYPE_STRING);// 日期
		setStringCellAndStyle(newsheet, baseinfo[2], 2, 90, null, Cell.CELL_TYPE_STRING);// 车型
		setStringCellAndStyle(newsheet, baseinfo[3], 48, 108, null, Cell.CELL_TYPE_STRING);// 批次

		setStringCellAndStyle(newsheet, baseinfo[4], 50, 72, null, Cell.CELL_TYPE_STRING);// 工位名称
		setStringCellAndStyle(newsheet, baseinfo[6], 51, 94, null, Cell.CELL_TYPE_STRING);// 工位编码

		setStringCellAndStyle(newsheet, page, 50, 107, null, Cell.CELL_TYPE_STRING);// 当前页码
		setStringCellAndStyle(newsheet, baseinfo[5], 50, 112, null, Cell.CELL_TYPE_STRING);// 总页码

		setStringCellAndStyle(newsheet, "新增", 49, 3, null, Cell.CELL_TYPE_STRING);// 标记
		setStringCellAndStyle(newsheet, "1", 49, 7, null, Cell.CELL_TYPE_STRING);// 处数
		setStringCellAndStyle(newsheet, username, 49, 23, null, Cell.CELL_TYPE_STRING);// 签字
		setStringCellAndStyle(newsheet, df2.format(new Date()), 49, 29, null, Cell.CELL_TYPE_STRING);// 日期

		int sheetindex = book.getSheetIndex(newsheet);
		book.setPrintArea(sheetindex, 0, 114, 0, 51);
		PrintSetup printSetup = newsheet.getPrintSetup();
		printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
		printSetup.setScale((short) 70);// 自定义缩放，此处100为无缩放
		printSetup.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)
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

	/*
	 * 从生成的报表中获取 “编制”、“日期”、“版次”信息
	 */
	private String[] getBaseinfomation() {
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
		String[] values = new String[7];
		XSSFSheet sheet = book.getSheetAt(0);
		XSSFRow row;
		XSSFCell cell;
		row = sheet.getRow(2);
		cell = row.getCell(6);
		values[0] = convertCellValueToString(cell);// 编制
		cell = row.getCell(30);
		values[1] = convertCellValueToString(cell);// 日期
		cell = row.getCell(90);
		values[2] = convertCellValueToString(cell);// 基本车型
		row = sheet.getRow(48);
		cell = row.getCell(108);
		values[3] = convertCellValueToString(cell);// 版次
		row = sheet.getRow(50);
		cell = row.getCell(72);
		values[4] = convertCellValueToString(cell);// 工位名称
		cell = row.getCell(112);
		String str = convertCellValueToString(cell);// 总页码
		values[5] = str;
		row = sheet.getRow(51);
		cell = row.getCell(94);
		values[6] = convertCellValueToString(cell);// 工序编号
		return values;
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
