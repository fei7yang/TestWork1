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
	SimpleDateFormat df2 = new SimpleDateFormat("yyyy.MM");// �������ڸ�ʽ

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

		// ������ʾ���Ȳ����ִ�в���
		ReportViwePanel viewPanel = new ReportViwePanel("���±���");
		viewPanel.setVisible(true);
		viewPanel.addInfomation("��ʼ���±���...\n", 20, 100);

		// ��ģ����ѡ���sheet���Ƶ�������
		TCComponentDataset dataset;
		if (model.equals("��ͨ��λģ��")) {
			dataset = FileUtil.getDatasetFile("DFL_Template_EngineeringWorkListStation");
			if (dataset == null) {
				viewPanel.addInfomation("����û���ҵ�������ҵ����ͨ��λģ�壬�������ģ��(����Ϊ��DFL_Template_EngineeringWorkListStation)\\n",
						100, 100);
				return;
			}
		} else if (model.equals("VIN����ģ��")) {
			dataset = FileUtil.getDatasetFile("DFL_Template_EngineeringWorkVINCarve");
			if (dataset == null) {
				viewPanel.addInfomation("����û���ҵ�������ҵ��VIN����ģ�壬�������ģ��(����Ϊ��DFL_Template_EngineeringWorkVINCarve)\\n", 100,
						100);
				return;
			}
		} else {
			dataset = FileUtil.getDatasetFile("DFL_Template_AdjustmentLine");
			if (dataset == null) {
				viewPanel.addInfomation("����û���ҵ�������ҵ�������ģ�壬�������ģ��(����Ϊ��DFL_Template_AdjustmentLine)\\n", 100, 100);
				return;
			}
		}
		String tempPath = getTempPath();
		File dirfile = new File(tempPath);
		if (!dirfile.exists()) {
			dirfile.mkdir();
		}
		// �������ݼ��ļ�
		File file = downloadFile((TCComponentDataset) dataset, tempPath);
		if (file == null) {
			System.out.println("�������ݼ��ļ�����");
			viewPanel.addInfomation("�������ݼ��ļ�����\n", 100, 100);
			return;
		}
		// �������ݼ��ļ�
		File factfile = downloadFile(factdatawet, tempPath);
		if (factfile == null) {
			System.out.println("�������ݼ��ļ�����");
			viewPanel.addInfomation("�������ݼ��ļ�����\n", 100, 100);
			return;
		}
		// ����VBS�ű�
		File scriptFile = Util.getRCPPluginInsideFile("CopySheet.vbs");
		if (scriptFile == null || !scriptFile.exists()) {
			viewPanel.addInfomation("����CopySheet.vbs�ű�����\n", 100, 100);
			// MessageBox.post("����SplitExcel.vbs�ű�����", "������ҵ����", MessageBox.WARNING);
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
		// ���ñ���vbs����sheet
		File aferfile = callVBSProgram(tempPath, factfile.getAbsolutePath(), vbsFilePath, modelsheetname);
		viewPanel.addInfomation("", 40, 100);
		if (aferfile == null) {
			viewPanel.addInfomation("����sheetʧ��\n", 100, 100);
			// MessageBox.post("����SplitExcel.vbs�ű�����", "������ҵ����", MessageBox.WARNING);
			return;
		}
		InputStream afterfilein = new FileInputStream(aferfile);
		book = NewOutputDataToExcel.creatXSSFWorkbook(afterfilein);
		// f���ƺ��sheet�ڵ�һҳ
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
			MessageBox.post("������sheet�����Ѵ��ڣ�sheet���Ʋ����ظ���", "��ʾ��Ϣ", MessageBox.INFORMATION);
			return ;
		}
		

		viewPanel.addInfomation("����д������...\n", 60, 100);

		// �����ɵı����л�ȡ �����ơ��������ڡ�������Ρ�����Ϣ��Ϣ
		String[] baseinfo = getBaseinfomation();

		// д���ļ�¼��Ϣ�ͻ�����Ϣ
		wirteDataToSheet(newsheet, baseinfo);

		// �޸���Чҳ
		updateValidPage();

		// �������ô�ӡ����
		setPrintArea();

		// ��ȡ��ҵ����
		int shs = book.getNumberOfSheets();
		String[] contents = new String[shs];
		for (int i = 0; i < shs; i++) {
			String sheetname = book.getSheetName(i);
			contents[i] = sheetname;
		}
		String procName = Util.getProperty(info.getMeDocument(), "object_name");
		String filename = Util.formatString(procName);
		NewOutputDataToExcel.exportFile(book, filename);
		// ������·
		{
			Util.callByPass(session, true);
		}
		// д����ҵ���ݺͺ�����Ҫ��
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
		// �ر���·
		{
			Util.callByPass(session, false);
		}

		// ִ����󣬰��ļ�ɾ��
		if (file.exists()) {
			file.delete();
		}
		if (factfile.exists()) {
			factfile.delete();
		}
		if (aferfile.exists()) {
			aferfile.delete();
		}

		viewPanel.addInfomation("���������ɣ����ں�װ������λ���󸽼��²鿴����...\n", 100, 100);
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
					System.out.println("vbs����sheet����");
				}

			} else {
				System.out.println("vbs����sheet����");
			}

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return null;
	}

	/**
	 * ͨ�����ƻ�ȡWordģ���ļ�
	 * 
	 * @param name
	 * @return ���ݼ�����
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
			printSetup.setScale((short) 70);// �Զ������ţ��˴�100Ϊ������
			printSetup.setLandscape(true); // ��ӡ����true������false������(Ĭ��)
		}
	}

	private void updateValidPage() {
		// TODO Auto-generated method stub
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // PSW����λ��
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("��Чҳ")) {
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
		// ����������ɫ
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
							setStringCellAndStyle(sheet, "��", i, 15 + 35 * col, style, Cell.CELL_TYPE_STRING);
						} else if (edition.equals("B")) {
							setStringCellAndStyle(sheet, "��", i, 19 + 35 * col, style, Cell.CELL_TYPE_STRING);
						} else if (edition.equals("C")) {
							setStringCellAndStyle(sheet, "��", i, 23 + 35 * col, style, Cell.CELL_TYPE_STRING);
						} else if (edition.equals("D")) {
							setStringCellAndStyle(sheet, "��", i, 27 + 35 * col, style, Cell.CELL_TYPE_STRING);
						} else if (edition.equals("E")) {
							setStringCellAndStyle(sheet, "��", i, 31 + 35 * col, style, Cell.CELL_TYPE_STRING);
						} else if (edition.equals("F")) {
							setStringCellAndStyle(sheet, "��", i, 35 + 35 * col, style, Cell.CELL_TYPE_STRING);
						} else {

						}
					}
				}

			}
		}

	}

	/*
	 * д���ļ�¼��Ϣ�ͻ�����Ϣ
	 */
	private void wirteDataToSheet(XSSFSheet newsheet, String[] baseinfo) throws TCException {
		// TODO Auto-generated method stub

		TCComponentUser user = session.getUser();
		String username = user.getUserName();

		// ҳ���һλΪ0������ʾ0������06A ��ʾΪ6A
		String page = "";
		if (sheetpages.substring(0, 1).equals("0")) {
			page = sheetpages.substring(1);
		} else {
			page = sheetpages;
		}

		setStringCellAndStyle(newsheet, baseinfo[0], 2, 6, null, Cell.CELL_TYPE_STRING); // ����
		setStringCellAndStyle(newsheet, baseinfo[1], 2, 30, null, Cell.CELL_TYPE_STRING);// ����
		setStringCellAndStyle(newsheet, baseinfo[2], 2, 90, null, Cell.CELL_TYPE_STRING);// ����
		setStringCellAndStyle(newsheet, baseinfo[3], 48, 108, null, Cell.CELL_TYPE_STRING);// ����

		setStringCellAndStyle(newsheet, baseinfo[4], 50, 72, null, Cell.CELL_TYPE_STRING);// ��λ����
		setStringCellAndStyle(newsheet, baseinfo[6], 51, 94, null, Cell.CELL_TYPE_STRING);// ��λ����

		setStringCellAndStyle(newsheet, page, 50, 107, null, Cell.CELL_TYPE_STRING);// ��ǰҳ��
		setStringCellAndStyle(newsheet, baseinfo[5], 50, 112, null, Cell.CELL_TYPE_STRING);// ��ҳ��

		setStringCellAndStyle(newsheet, "����", 49, 3, null, Cell.CELL_TYPE_STRING);// ���
		setStringCellAndStyle(newsheet, "1", 49, 7, null, Cell.CELL_TYPE_STRING);// ����
		setStringCellAndStyle(newsheet, username, 49, 23, null, Cell.CELL_TYPE_STRING);// ǩ��
		setStringCellAndStyle(newsheet, df2.format(new Date()), 49, 29, null, Cell.CELL_TYPE_STRING);// ����

		int sheetindex = book.getSheetIndex(newsheet);
		book.setPrintArea(sheetindex, 0, 114, 0, 51);
		PrintSetup printSetup = newsheet.getPrintSetup();
		printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
		printSetup.setScale((short) 70);// �Զ������ţ��˴�100Ϊ������
		printSetup.setLandscape(true); // ��ӡ����true������false������(Ĭ��)
	}

	// �Ե�Ԫ��ֵ
	public static void setStringCellAndStyle(XSSFSheet sheet, String value, int rowIndex, int cellIndex,
			XSSFCellStyle Style, int celltype) {

		// �����������ַ��͵����� 10Ϊ���ͣ�11Ϊdouble��

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
	 * �����ɵı����л�ȡ �����ơ��������ڡ�������Ρ���Ϣ
	 */
	private String[] getBaseinfomation() {
		// TODO Auto-generated method stub
		XSSFCellStyle cellStyle2 = book.createCellStyle();
		Font font2 = book.createFont();
		font2.setColor(IndexedColors.BLUE.getIndex());
		font2.setBoldweight(Font.BOLDWEIGHT_BOLD); // ����Ӵ�
		font2.setFontHeightInPoints((short) 16);
		font2.setFontName("����");
		cellStyle2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����
		cellStyle2.setAlignment(XSSFCellStyle.ALIGN_LEFT);
		cellStyle2.setFont(font2);
		String[] values = new String[7];
		XSSFSheet sheet = book.getSheetAt(0);
		XSSFRow row;
		XSSFCell cell;
		row = sheet.getRow(2);
		cell = row.getCell(6);
		values[0] = convertCellValueToString(cell);// ����
		cell = row.getCell(30);
		values[1] = convertCellValueToString(cell);// ����
		cell = row.getCell(90);
		values[2] = convertCellValueToString(cell);// ��������
		row = sheet.getRow(48);
		cell = row.getCell(108);
		values[3] = convertCellValueToString(cell);// ���
		row = sheet.getRow(50);
		cell = row.getCell(72);
		values[4] = convertCellValueToString(cell);// ��λ����
		cell = row.getCell(112);
		String str = convertCellValueToString(cell);// ��ҳ��
		values[5] = str;
		row = sheet.getRow(51);
		cell = row.getCell(94);
		values[6] = convertCellValueToString(cell);// ������
		return values;
	}

	private static String convertCellValueToString(Cell cell) {
		if (cell == null) {
			return null;
		}
		String returnValue = null;
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC: // ����
			Double doubleValue = cell.getNumericCellValue();
			// ��ʽ����ѧ��������ȡһλ����
			DecimalFormat df = new DecimalFormat("0");
			returnValue = df.format(doubleValue);
			break;
		case Cell.CELL_TYPE_STRING: // �ַ���
			returnValue = cell.getStringCellValue();
			break;
		case Cell.CELL_TYPE_BOOLEAN: // ����
			Boolean booleanValue = cell.getBooleanCellValue();
			returnValue = booleanValue.toString();
			break;
		case Cell.CELL_TYPE_BLANK: // ��ֵ
			break;
		case Cell.CELL_TYPE_FORMULA: // ��ʽ
			returnValue = cell.getCellFormula();
			break;
		case Cell.CELL_TYPE_ERROR: // ����
			break;
		default:
			break;
		}
		return returnValue;
	}
}
