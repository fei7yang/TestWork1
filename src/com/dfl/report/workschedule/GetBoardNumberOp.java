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
	private List<WeldPointBoardInformation> baseinfolist;// ������Ϣ�������

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
		// ������ʾ���Ȳ����ִ�в���
		ReportViwePanel viewPanel = new ReportViwePanel("���±���");
		viewPanel.setVisible(true);
		viewPanel.addInfomation("��ʼ���±���...\n", 20, 100);

		TCComponentUser user = session.getUser();
		String username = user.getUserName();

		// ��ȡ������Ϣ
//		String baseName = "222.������Ϣ";
//		baseinfolist = getBaseinfomation(topbomline.window().getTopBOMLine(), baseName);

		viewPanel.addInfomation("���ڸ��±���...\n", 30, 100);

		// ��ȡ�㺸sheetҳ
		List sheetlist = getSpotWeldingSheets();

		viewPanel.addInfomation("", 50, 100);

		// ���ݺ����ţ���ȡ��Ӧ�İ����Ų�д��
		writeBoradDataTOSheet(sheetlist);

		viewPanel.addInfomation("", 60, 100);

		String procName = Util.getProperty(info.getMeDocument(), "object_name");
		String filename = Util.formatString(procName);
		NewOutputDataToExcel.exportFile(book, filename);

		// ������·
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
		// �ر���·
		{
			Util.callByPass(session, false);
		}

		viewPanel.addInfomation("���������ɣ����ں�װ������λ���󸽼��²鿴����...\n", 100, 100);
	}

	private void writeBoradDataTOSheet(List sheetlist) {
		// TODO Auto-generated method stub
		if (sheetlist != null && sheetlist.size() > 0) {
			// ����������ɫ
			Font font = book.createFont();
			font.setColor((short) 12);// ��ɫ����
			font.setFontName("MS PGothic");
			font.setFontHeightInPoints((short) 12);
			XSSFCellStyle style = book.createCellStyle();
			style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
			style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
			style.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
			// style.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
			style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style.setFont(font);

			XSSFCellStyle style2 = book.createCellStyle();
			style2.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
			// style2.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
			style2.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
			// style2.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
			style2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style2.setFont(font);

			XSSFCellStyle style3 = book.createCellStyle();
			style3.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
			// style3.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
			style3.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
			style3.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
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
	 * ���ݺ���Ų��Ұ�����
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
	 * ��ȡ�㺸sheet
	 */
	private List getSpotWeldingSheets() {
		// TODO Auto-generated method stub
		int sheetnum = 0;
		List sheetList = new ArrayList();
		sheetnum = book.getNumberOfSheets();
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("�㺸")) {
				sheetList.add(sheetname);
			}
		}
		return sheetList;
	}

	/*
	 * ��ȡ������Ϣ����Ϣ
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
			case Cell.CELL_TYPE_NUMERIC: // ����
				Double doubleValue = cell.getNumericCellValue();
				// ��ʽ����ѧ��������ȡһλ����
				DecimalFormat df = new DecimalFormat("0.0");
				returnValue = df.format(doubleValue);
				break;
			case Cell.CELL_TYPE_STRING: // �ַ���
				cell.setCellType(Cell.CELL_TYPE_STRING);
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
		}
		return returnValue;
	}

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

		//cell.setCellStyle(Style);

	}
}
