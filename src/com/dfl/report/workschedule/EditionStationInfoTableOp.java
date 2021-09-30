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
	SimpleDateFormat df2 = new SimpleDateFormat("yyyy.MM");// �������ڸ�ʽ
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

		// ������ʾ���Ȳ����ִ�в���
		ReportViwePanel viewPanel = new ReportViwePanel("���±���");
		viewPanel.setVisible(true);
		viewPanel.addInfomation("��ʼ���±���...\n", 20, 100);

		TCComponentUser user = session.getUser();
		String username = user.getUserName();
		String date = df2.format(new Date());
		TCComponentGroup group = session.getGroup();
		// ����
		String groupname = group.getLocalizedFullName();
		// ���п�
		String department = "";
		if (groupname != null && (groupname.contains("ͬ�ڹ��̿�") || groupname.contains("simultaneous Engineering Section"))) {
			department = "H30";
		} else if (groupname != null && (groupname.contains("��װ������") || groupname.contains("Body Assembly Engineering Section"))) {
			department = "VE2";
		} else {
			department = "VE2";
		}

		viewPanel.addInfomation("���ڸ��±���...\n", 30, 100);
		// ��Чҳ�������
		dealValidPage();

		// ���ı������������ơ��������ڡ��������
		dealClearAndwriteDateToSheet(username, date,department,Edition);

		viewPanel.addInfomation("", 50, 100);

		// ҳ������
		dealPageRearrangement();

		viewPanel.addInfomation("", 60, 100);

		// ������·
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
		// �ر���·
		{
			Util.callByPass(session, false);
		}

		viewPanel.addInfomation("���������ɣ����ں�װ������λ���󸽼��²鿴����...\n", 100, 100);
	}

	/*
	 * ��������ֵ
	 */
	private void setPropertyValue(TCComponent tcc, String property, String value) throws TCException {
		TCProperty p = tcc.getTCProperty(property);
		if (p != null) {
			p.setStringValue(value);
		}
	}

	/*
	 * ҳ������
	 */
	private void dealPageRearrangement() {
		// TODO Auto-generated method stub
		int sheetnum = book.getNumberOfSheets();
		// ����Ƚ�����,��ʼֵΪ��һ��sheet���ƣ����������ͬ������Ҫ�����ƺ�������1,2......
		String tempname = "";
		String sheetAllname;
		int num = 1;
		Pattern p = Pattern.compile("[0-9a-fA-F]"); // ��������ĸ
		Pattern p2 = Pattern.compile("[0-9]"); // ��������ĸ
		// �Ȱ���˳���������������չ����������������ظ����
		for (int i = 0; i < sheetnum; i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			String sheetname = sheet.getSheetName();
			sheetAllname = sheetname + (i + 1);
			book.setSheetName(i, sheetAllname);
		}
		for (int i = 0; i < sheetnum; i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			// ��ȡǰ��λ��ȥ����ĸ
			String sheetname = "";
			String oldsheetname = sheet.getSheetName();
			if (oldsheetname.length() > 2) {
				Matcher m = p.matcher(oldsheetname.substring(0, 3));
				Matcher m2 = p2.matcher(oldsheetname.substring(3));
				sheetname = m.replaceAll("").trim() + m2.replaceAll("").trim();
			} else {
				sheetname = oldsheetname;
			}
			// ��һ���Ȳ��Ƚϣ��ӵڶ�����ʼ�͵�һ���Ƚ�
			if (i == 0) {
				tempname = sheetname;
				sheetAllname = String.format("%02d", i + 1) + sheetname;
				book.setSheetName(i, sheetAllname);

			} else {
				if (sheetname.contains(tempname)) {
					// ���numΪ1����˵��sheetͬ���ĵ�һ����Ҫ�������������ֺ�׺
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
			// ���ô�ӡ����
			book.removePrintArea(i);
		}

		for (int i = 0; i < sheetnum; i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			book.setPrintArea(i, 0, 114, 0, 51);
			PrintSetup printSetup = sheet.getPrintSetup();
			printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
			printSetup.setScale((short) 70);// �Զ������ţ��˴�100Ϊ������
			printSetup.setLandscape(true); // ��ӡ����true������false������(Ĭ��)
		}
	}

	/*
	 * ���ı������������ơ��������ڡ��������
	 */
	private void dealClearAndwriteDateToSheet(String username, String date, String department, String edition) {
		// TODO Auto-generated method stub
		// ��������
//		Font font = book.createFont();
//		font.setColor((short) 12);
//		font.setFontName("����");
//		font.setFontHeightInPoints((short) 16);
		// ����һ����ʽ
		XSSFCellStyle cellStyle1 = null;
//		cellStyle1.setBorderBottom(XSSFCellStyle.BORDER_THIN);
//		cellStyle1.setBorderLeft(XSSFCellStyle.BORDER_THIN);
//		cellStyle1.setBorderRight(XSSFCellStyle.BORDER_THIN);
//		cellStyle1.setBorderTop(XSSFCellStyle.BORDER_THIN);
//		cellStyle1.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����
//		cellStyle1.setAlignment(XSSFCellStyle.ALIGN_CENTER);
//		cellStyle1.setFont(font);

		XSSFCellStyle cellStyle2 = null;
//		Font font2 = book.createFont();
//		font2.setColor(IndexedColors.BLUE.getIndex());
//		font2.setBoldweight(Font.BOLDWEIGHT_BOLD); // ����Ӵ�
//		font2.setFontHeightInPoints((short) 16);
//		font2.setFontName("����");
//		cellStyle2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����
//		cellStyle2.setAlignment(XSSFCellStyle.ALIGN_LEFT);
//		cellStyle2.setFont(font2);

		// ��������
//		Font font3 = book.createFont();
//		font3.setColor((short) 12);
//		font3.setFontName("����");
//		font3.setFontHeightInPoints((short) 12);
		// ����һ����ʽ
		XSSFCellStyle cellStyle3 = null;
//		cellStyle3.setBorderBottom(XSSFCellStyle.BORDER_THIN);
//		cellStyle3.setBorderLeft(XSSFCellStyle.BORDER_THIN);
//		cellStyle3.setBorderRight(XSSFCellStyle.BORDER_THIN);
//		cellStyle3.setBorderTop(XSSFCellStyle.BORDER_THIN);
//		cellStyle3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����
//		cellStyle3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
//		cellStyle3.setFont(font3);

		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		for (int i = 0; i < sheetnum; i++) {
			XSSFSheet sheet = book.getSheetAt(i);

			setStringCellAndStyle(sheet, department, 2, 0, cellStyle1, Cell.CELL_TYPE_STRING); // ���п�
			setStringCellAndStyle(sheet, username, 2, 6, cellStyle1, Cell.CELL_TYPE_STRING); // ����
			setStringCellAndStyle(sheet, date, 2, 30, cellStyle1, Cell.CELL_TYPE_STRING);// ����
			setStringCellAndStyle(sheet, Integer.toString(i + 1), 50, 107, cellStyle2, 10);// ��ǰҳ��
			setStringCellAndStyle(sheet, Integer.toString(sheetnum), 50, 112, cellStyle2, 10);// ��ҳ��
//			setStringCellAndStyle(sheet, edition, 48, 108, cellStyle2, 10);// ����

			setStringCellAndStyle(sheet, "", 49, 3, cellStyle3, Cell.CELL_TYPE_STRING);// ���
			setStringCellAndStyle(sheet, "", 49, 7, cellStyle3, Cell.CELL_TYPE_STRING);// ����
			setStringCellAndStyle(sheet, "", 49, 23, cellStyle3, Cell.CELL_TYPE_STRING);// ǩ��
			setStringCellAndStyle(sheet, "", 49, 29, cellStyle3, Cell.CELL_TYPE_STRING);// ����

			setStringCellAndStyle(sheet, "", 50, 3, cellStyle3, Cell.CELL_TYPE_STRING);// ���
			setStringCellAndStyle(sheet, "", 50, 7, cellStyle3, Cell.CELL_TYPE_STRING);// ����
			setStringCellAndStyle(sheet, "", 50, 23, cellStyle3, Cell.CELL_TYPE_STRING);// ǩ��
			setStringCellAndStyle(sheet, "", 50, 29, cellStyle3, Cell.CELL_TYPE_STRING);// ����

			setStringCellAndStyle(sheet, "", 51, 3, cellStyle3, Cell.CELL_TYPE_STRING);// ���
			setStringCellAndStyle(sheet, "", 51, 7, cellStyle3, Cell.CELL_TYPE_STRING);// ����
			setStringCellAndStyle(sheet, "", 51, 23, cellStyle3, Cell.CELL_TYPE_STRING);// ǩ��
			setStringCellAndStyle(sheet, "", 51, 29, cellStyle3, Cell.CELL_TYPE_STRING);// ����

			setStringCellAndStyle(sheet, "", 49, 35, cellStyle3, Cell.CELL_TYPE_STRING);// ���
			setStringCellAndStyle(sheet, "", 49, 39, cellStyle3, Cell.CELL_TYPE_STRING);// ����
			setStringCellAndStyle(sheet, "", 49, 55, cellStyle3, Cell.CELL_TYPE_STRING);// ǩ��
			setStringCellAndStyle(sheet, "", 49, 61, cellStyle3, Cell.CELL_TYPE_STRING);// ����

			setStringCellAndStyle(sheet, "", 50, 35, cellStyle3, Cell.CELL_TYPE_STRING);// ���
			setStringCellAndStyle(sheet, "", 50, 39, cellStyle3, Cell.CELL_TYPE_STRING);// ����
			setStringCellAndStyle(sheet, "", 50, 55, cellStyle3, Cell.CELL_TYPE_STRING);// ǩ��
			setStringCellAndStyle(sheet, "", 50, 61, cellStyle3, Cell.CELL_TYPE_STRING);// ����

			setStringCellAndStyle(sheet, "", 51, 35, cellStyle3, Cell.CELL_TYPE_STRING);// ���
			setStringCellAndStyle(sheet, "", 51, 39, cellStyle3, Cell.CELL_TYPE_STRING);// ����
			setStringCellAndStyle(sheet, "", 51, 55, cellStyle3, Cell.CELL_TYPE_STRING);// ǩ��
			setStringCellAndStyle(sheet, "", 51, 61, cellStyle3, Cell.CELL_TYPE_STRING);// ����

		}
	}

	/*
	 * ��Чҳ�������
	 */
	private void dealValidPage() {
		// TODO Auto-generated method stub
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; //
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
		XSSFSheet sheet = book.getSheetAt(sheetAtIndex);
		// ����������ɫ
//		Font font = book.createFont();
//		font.setColor((short) 12);// ��ɫ����
//		font.setFontName("����");
//		font.setFontHeightInPoints((short) 14);
		XSSFCellStyle style = null;
//		style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
//		style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
//		style.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
//		style.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
//		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
//		style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
//		style.setFont(font);
		// �����
		int page = 3;
		for (int i = 0; i < page; i++) {
			for (int j = 0; j < 40; j++) {
				for (int k = 0; k < 7; k++) {
					setStringCellAndStyle(sheet, "", 7 + j, 11 + 35 * i + k * 4, style, Cell.CELL_TYPE_STRING); //
				}
			}
		}
		// ������д
		page = (sheetnum - 1) / 40 + 1;
		for (int i = 0; i < page; i++) {
			if (i == page - 1) {
				for (int j = 0; j < sheetnum - 40 * i; j++) {
					setStringCellAndStyle(sheet, "��", 7 + j, 11 + 35 * i, style, Cell.CELL_TYPE_STRING); // ����
				}
			} else {
				for (int j = 0; j < 40; j++) {
					setStringCellAndStyle(sheet, "��", 7 + j, 11 + 35 * i, style, Cell.CELL_TYPE_STRING); // ����
				}
			}
		}
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

}
