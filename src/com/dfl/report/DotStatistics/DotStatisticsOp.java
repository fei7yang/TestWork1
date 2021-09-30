package com.dfl.report.DotStatistics;

import java.io.File;
import java.io.InputStream;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dfl.report.util.FileUtil;
import com.dfl.report.util.NewOutputDataToExcel;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentItemType;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;

public class DotStatisticsOp {

	private TCSession session;
	private InterfaceAIFComponent[] aifComponents;
	private TCComponent savefolder;
	private ArrayList other = new ArrayList();
	private int allsum = 0;// �ܺ�����
	private int rswsum = 0;// �Զ�������
	private Map<String, Integer[]> statistics = new HashMap<String, Integer[]>();// �˹���������ͳ�Ƽ���
	SimpleDateFormat df = new SimpleDateFormat("yyyyMMddHH");// �������ڸ�ʽ

	public DotStatisticsOp(TCSession session, InterfaceAIFComponent[] aifComponents, TCComponent savefolder)
			throws TCException {
		// TODO Auto-generated constructor stub
		this.session = session;
		this.aifComponents = aifComponents;
		this.savefolder = savefolder;
		initUI();
	}

	private void initUI() throws TCException {
		// TODO Auto-generated method stub

		// ��ʾ�����������
		// ������ʾ���Ȳ����ִ�в���
		ReportViwePanel viewPanel = new ReportViwePanel("���ɱ���");
		viewPanel.setVisible(true);
		TCComponentBOMLine topbl = (TCComponentBOMLine) aifComponents[0];

		viewPanel.addInfomation("���ڻ�ȡģ��...\n", 10, 100);
		// ��ѯĿ¼����ģ��
		InputStream inputStream = FileUtil.getTemplateFile("DFL_Template_DotStatistics");

		if (inputStream == null) {
			viewPanel.addInfomation("����û���ҵ����ͳ�Ʊ�ģ�壬�������ģ��(����Ϊ��DFL_Template_DotStatistics)\n", 100, 100);
			return;
		}
		String FamlilyCode = Util.getProperty(topbl, "bl_rev_project_ids");// ��������
		String VehicleNo = Util.getDFLProjectIdVehicle(FamlilyCode);
		if(VehicleNo== null || VehicleNo.isEmpty()) {
			VehicleNo = FamlilyCode;
		}
		// ��������  ��Ϊ��BOP������ȡֵ
		String factoryname = "";
		String bopname = Util.getProperty(topbl, "bl_rev_object_name");
		String[] bopnames = bopname.split("_");
		if(bopnames.length>2) {
			String factory = bopnames[2];
			if(factory.length()>2) {
				factoryname = factory.substring(0, 3);
			}
		}
		// ��ȡ������BOe
//		TCComponent[] boelist = topbl.getItemRevision().getRelatedComponents("IMAN_MEWorkArea");
//		if (boelist != null && boelist.length > 0) {
//			TCComponentItemRevision boerev = (TCComponentItemRevision) boelist[0];
//			factoryname = factoryname + Util.getProperty(boerev, "object_name");
//		}
		viewPanel.addInfomation("", 20, 100);
		// ����BOP���㣬��ȡ���е������ߣ��������û�з��������ߣ���Ϊ����
		ArrayList asahi = getAsahiLine(topbl);

		viewPanel.addInfomation("��ʼ�������...\n", 40, 100);

		// ͳ�Ƶ㺸�����µĺ�����
		Map<String, ArrayList> map = getAllWeldnumInfo(asahi);

		String[] str = new String[6];
		NumberFormat numberFormat = NumberFormat.getInstance();
		String result = "";
		if (allsum == 0) {
			result = "0";
		} else {
			result = numberFormat.format((float) rswsum / (float) allsum * 100);
		}
		str[0] = factoryname;
		str[1] = VehicleNo;
		str[2] = result + "%";
		str[3] = Integer.toString(allsum);
		str[4] = Integer.toString(rswsum);

		viewPanel.addInfomation("��ʼд���ݣ������ĵȴ�...\n", 60, 100);

		XSSFWorkbook book = NewOutputDataToExcel.creatXSSFWorkbook(inputStream);
		// д������
		writeDataToSheet(book, statistics, str, map, asahi);
				
		String sheetname = 	VehicleNo + "-" + "���ͳ����ϸҳ";
		book.setSheetName(1, sheetname);
			

		String date = df.format(new Date());
		String datasetname = VehicleNo  + "_���ͳ�Ʊ�" + "_" + date + "ʱ";
		String filename = Util.formatString(datasetname);

		NewOutputDataToExcel.exportFile(book, filename);
		
		
		viewPanel.addInfomation("", 80, 100);

		// NewOutputDataToExcel.openFile(FileUtil.getReportFileName(filename.trim()));
		saveFiles(filename, datasetname, savefolder, session);

		viewPanel.addInfomation("���������ɣ�����ѡ�񱣴���ļ����²鿴��", 100, 100);
	}

	/*
	 * ****************************************************** �����ɵı�������ָ�����ļ�����
	 */
	public void saveFiles(String filename, String datasetName, TCComponent folder, TCSession session) {
		try {
			String fullFileName = FileUtil.getReportFileName(filename);
			TCComponentFolder savefolder = (TCComponentFolder) folder;
			TCComponentItemType tcccomponentitemtype = (TCComponentItemType) session.getTypeComponent("B8_BIWProcDoc");
			TCComponentItem tccomponentitem = tcccomponentitemtype.create("", "", "B8_BIWProcDoc", datasetName, "desc",
					null);
			tccomponentitem.setProperty("b8_BIWProcDocType", "AO");
			tccomponentitem.lock();
			tccomponentitem.save();
			tccomponentitem.unlock();
			TCComponentItemRevision rev = tccomponentitem.getLatestItemRevision();
			TCComponentDataset ds = Util.createDataset(session, datasetName, fullFileName, "MSExcelX", "excel");

			rev.add("IMAN_specification", ds);

			// ����ĵ������ݼ��Ĺ�ϵ
			savefolder.add("contents", tccomponentitem);

			// ɾ���м��ļ�
			File file = new File(fullFileName);
			if (file.isFile()) {
				file.delete();
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// д������
	private void writeDataToSheet(XSSFWorkbook book, Map<String, Integer[]> statistics2, String[] str,
			Map<String, ArrayList> map, ArrayList asahi) {
		// TODO Auto-generated method stub

		// ����������ɫ
		Font font = book.createFont();
		// font.setColor((short) 12);// ��ɫ����
		font.setFontHeightInPoints((short) 11);
		XSSFCellStyle style = book.createCellStyle();
		style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
		style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
		style.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
		style.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style.setFont(font);
		Font font2 = book.createFont();
		font2.setColor((short) 2);// ��ɫ����
		font2.setFontHeightInPoints((short) 11);
		XSSFCellStyle style2 = book.createCellStyle();
		style2.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
		style2.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
		style2.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
		style2.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
		style2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style2.setFont(font2);
		// ��ɫ����
		XSSFCellStyle style3 = book.createCellStyle();
		style3.setFillForegroundColor(IndexedColors.VIOLET.getIndex());
		style3.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style3.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
		style3.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
		style3.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
		style3.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
		style3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style3.setFont(font);
		// ��ɫ����
		XSSFCellStyle style4 = book.createCellStyle();
		style4.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
		style4.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style4.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
		style4.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
		style4.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
		style4.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
		style4.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style4.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style4.setFont(font);
		// ��д���һ��sheet����
		XSSFSheet sheet1 = book.getSheetAt(0);
		// ����-�Զ�������
		setStringCellAndStyle(sheet1, str[0], 1, 1, style, Cell.CELL_TYPE_STRING);
		setStringCellAndStyle(sheet1, str[1], 1, 2, style, Cell.CELL_TYPE_STRING);
		setStringCellAndStyle(sheet1, str[2], 1, 3, style, Cell.CELL_TYPE_STRING);
		setStringCellAndStyle(sheet1, str[3], 1, 4, style, 10);
		setStringCellAndStyle(sheet1, str[4], 1, 5, style2, 10);
		
		//�����Զ��п�
        for (int i = 0; i < 10; i++) {
            sheet1.autoSizeColumn(i);
        }


		// ��д�ڶ���sheet
		XSSFSheet sheet2 = book.getSheetAt(1);
		// ��̬���ر�ͷ
		int colnum = 0;
		int colnum2 = 0;
		CellRangeAddress region;
		for (int i = 0; i < asahi.size(); i++) {
			TCComponentBOMLine bl = (TCComponentBOMLine) asahi.get(i);
			String templinename = Util.getProperty(bl, "bl_rev_object_name");
			String linename = templinename.replaceAll("\\d+", "").replace("-", "").replace("_", "");;// ȥ�����֣�����������
			if (statistics2.containsKey(linename)) {
				Integer[] instr = statistics2.get(linename);
				setStringCellAndStyle(sheet1, linename + "��������", 0, 7 + 3 * colnum, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet1, linename + "�˹�", 0, 8 + 3 * colnum, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet1, linename + "������", 0, 9 + 3 * colnum, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet1, instr[0].toString(), 1, 7 + 3 * colnum, style, 10);
				setStringCellAndStyle(sheet1, instr[1].toString(), 1, 8 + 3 * colnum, style, 10);
				setStringCellAndStyle(sheet1, instr[2].toString(), 1, 9 + 3 * colnum, style2, 10);
				colnum++;
			}
			if (sheet2 != null) {
				if (map.containsKey(linename)) {
					setStringCellAndStyle(sheet2, linename, 0, 0 + 4 * colnum2, style4, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet2, "RH", 0, 1 + 4 * colnum2, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet2, "LH", 0, 2 + 4 * colnum2, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet2, "", 0, 3 + 4 * colnum2, style3, Cell.CELL_TYPE_STRING);

					ArrayList list = map.get(linename);
					String begin = "";
					int beginrow = 1;// ��ʼ��
					int endrow = 0;// ��ֹ��
					int n = 0;// �����
					for (int j = 0; j < list.size(); j++) {
						String[] values = (String[]) list.get(j);
						if (values[0].equals("�ܼ�")) {
							setStringCellAndStyle(sheet2, values[0], 1 + j, 0 + 4 * colnum2, style2,
									Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sheet2, values[1], 1 + j, 1 + 4 * colnum2, style2, 10);
							setStringCellAndStyle(sheet2, values[2], 1 + j, 2 + 4 * colnum2, style2, 10);
						} else {
							setStringCellAndStyle(sheet2, values[0], 1 + j, 0 + 4 * colnum2, style,
									Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sheet2, values[1], 1 + j, 1 + 4 * colnum2, style, 10);
							setStringCellAndStyle(sheet2, values[2], 1 + j, 2 + 4 * colnum2, style, 10);
						}
						setStringCellAndStyle(sheet2, "", 1 + j, 3 + 4 * colnum2, style3, Cell.CELL_TYPE_STRING);

						if (j == 0) {
							begin = values[0];
						} else {
							if (!begin.equals(values[0])) {
								endrow = beginrow + n;
								region = new CellRangeAddress(beginrow, endrow, (short) (0 + 4 * colnum2),
										(short) (0 + 4 * colnum2));
								sheet2.addMergedRegion(region);
								begin = values[0];
								beginrow = endrow + 1;
								n = 0;
							} else {
								n++;
							}
						}
						if (j == list.size() - 1) {
							endrow = beginrow + n;
							region = new CellRangeAddress(beginrow, endrow, (short) (0 + 4 * colnum2),
									(short) (0 + 4 * colnum2));
							sheet2.addMergedRegion(region);
						}

					}
					colnum2++;
				}

			}
		}
		if (statistics2.containsKey("����")) {
			Integer[] instr = statistics2.get("����");
			setStringCellAndStyle(sheet1, "����" + "��������", 0, 7 + 3 * colnum, style, Cell.CELL_TYPE_STRING);
			setStringCellAndStyle(sheet1, "����" + "�˹�", 0, 8 + 3 * colnum, style, Cell.CELL_TYPE_STRING);
			setStringCellAndStyle(sheet1, "����" + "������", 0, 9 + 3 * colnum, style, Cell.CELL_TYPE_STRING);
			setStringCellAndStyle(sheet1, instr[0].toString(), 1, 7 + 3 * colnum, style, 10);
			setStringCellAndStyle(sheet1, instr[1].toString(), 1, 8 + 3 * colnum, style, 10);
			setStringCellAndStyle(sheet1, instr[2].toString(), 1, 9 + 3 * colnum, style2, 10);

		}
		if (map.containsKey("����")) {
			setStringCellAndStyle(sheet2, "����", 0, 0 + 4 * colnum2, style4, Cell.CELL_TYPE_STRING);
			setStringCellAndStyle(sheet2, "RH", 0, 1 + 4 * colnum2, style, Cell.CELL_TYPE_STRING);
			setStringCellAndStyle(sheet2, "LH", 0, 2 + 4 * colnum2, style, Cell.CELL_TYPE_STRING);
			setStringCellAndStyle(sheet2, "", 0, 3 + 4 * colnum2, style3, Cell.CELL_TYPE_STRING);

			ArrayList list = map.get("����");
			String begin = "";
			int beginrow = 1;// ��ʼ��
			int endrow = 0;// ��ֹ��
			int n = 0;// �����

			for (int j = 0; j < list.size(); j++) {
				String[] values = (String[]) list.get(j);
				if (values[0].equals("�ܼ�")) {
					setStringCellAndStyle(sheet2, values[0], 1 + j, 0 + 4 * colnum2, style2, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet2, values[1], 1 + j, 1 + 4 * colnum2, style2, 10);
					setStringCellAndStyle(sheet2, values[2], 1 + j, 2 + 4 * colnum2, style2, 10);
				} else {
					setStringCellAndStyle(sheet2, values[0], 1 + j, 0 + 4 * colnum2, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet2, values[1], 1 + j, 1 + 4 * colnum2, style, 10);
					setStringCellAndStyle(sheet2, values[2], 1 + j, 2 + 4 * colnum2, style, 10);
				}
				setStringCellAndStyle(sheet2, "", 1 + j, 3 + 4 * colnum2, style3, Cell.CELL_TYPE_STRING);

				if (j == 0) {
					begin = values[0];
				} else {
					if (!begin.equals(values[0])) {
						endrow = beginrow + n;
						region = new CellRangeAddress(beginrow, endrow, (short) (0 + 4 * colnum2),
								(short) (0 + 4 * colnum2));
						sheet2.addMergedRegion(region);
						begin = values[0];
						beginrow = endrow + 1;
						n = 0;
					} else {
						n++;
					}
				}
				if (j == list.size() - 1) {
					endrow = beginrow + n;
					region = new CellRangeAddress(beginrow, endrow, (short) (0 + 4 * colnum2),
							(short) (0 + 4 * colnum2));
					sheet2.addMergedRegion(region);
				}
			}
		}
//		//�����Զ��п�
//        for (int i = 0; i < 3*colnum; i++) {
//            sheet1.autoSizeColumn(i);
//        }
//        // �������Ĳ����Զ������п������
//        this.setSizeColumn(sheet1, 3*colnum);
//        for (int i = 0; i < 3*colnum2; i++) {
//            sheet2.autoSizeColumn(i);
//        }
//        this.setSizeColumn(sheet2, 3*colnum2);
	}
	// ����Ӧ���(����֧��)
    private void setSizeColumn(XSSFSheet sheet, int size) {
        for (int columnNum = 0; columnNum < size; columnNum++) {
            int columnWidth = sheet.getColumnWidth(columnNum) / 256;
            for (int rowNum = 0; rowNum < sheet.getLastRowNum(); rowNum++) {
                XSSFRow currentRow;
                //��ǰ��δ��ʹ�ù�
                if (sheet.getRow(rowNum) == null) {
                    currentRow = sheet.createRow(rowNum);
                } else {
                    currentRow = sheet.getRow(rowNum);
                }
 
                if (currentRow.getCell(columnNum) != null) {
                    XSSFCell currentCell = currentRow.getCell(columnNum);
                    if (currentCell.getCellType() == XSSFCell.CELL_TYPE_STRING) {
                        int length = currentCell.getStringCellValue().getBytes().length;
                        if (columnWidth < length) {
                            columnWidth = length;
                        }
                    }
                }
            }
            sheet.setColumnWidth(columnNum, columnWidth * 256);
        }
    }
	public void setStringCellAndStyle(XSSFSheet sheet, String value, int rowIndex, int cellIndex, XSSFCellStyle Style,
			int celltype) {

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

		cell.setCellStyle(Style);

	}

	// ͳ�Ƶ㺸�����µĺ�����
	private Map<String, ArrayList> getAllWeldnumInfo(ArrayList asahi) throws TCException {
		// TODO Auto-generated method stub
		Map<String, ArrayList> map = new HashMap<String, ArrayList>();
		if (asahi != null && asahi.size() > 0) {
			for (int i = 0; i < asahi.size(); i++) {
				ArrayList dataList = new ArrayList();
				int psw = 0;
				int rsw = 0;
				TCComponentBOMLine bl = (TCComponentBOMLine) asahi.get(i);
				String templinename = Util.getProperty(bl, "bl_rev_object_name");
				String linename = templinename.replaceAll("\\d+", "").replace("-", "").replace("_", "");// ȥ�����֣�����������
				ArrayList linelist = Util.getChildrenByBOMLine(bl, "B8_BIWMEProcLineRevision");// ʵ�ʺ�װ����
				if (linelist != null && linelist.size() > 0) {
					for (int j = 0; j < linelist.size(); j++) {
						boolean flag = false;// �ж�ʵ�ʲ����Ƿ��й�λ��û�в�����ܼ�
						TCComponentBOMLine linebl = (TCComponentBOMLine) linelist.get(j);
						int rhsum = 0;// RH �ܼ�����
						int lhsum = 0;// LH �ܼ�����
						boolean flag2 = Util.getIsMEProcStat(linebl);
						String lname = Util.getProperty(linebl.getItemRevision(), "b8_LineType") + Util.getProperty(linebl, "bl_rev_object_name"); // ʵ�ʲ������� ��Ҫƴ������ʽ�����   20191202�޸�
						ArrayList statelist = Util.getChildrenByBOMLine(linebl, "B8_BIWMEProcStatRevision");// ��װ��λ����
						if (statelist != null && statelist.size() > 0) {
							flag = true;
							for (int k = 0; k < statelist.size(); k++) {
								TCComponentBOMLine statebl = (TCComponentBOMLine) statelist.get(k);
								String statename = Util.getProperty(statebl, "bl_rev_object_name"); // ��λ����
								ArrayList dislist = Util.getChildrenByBOMLine(statebl, "B8_BIWDiscreteOPRevision");// �㺸����
								if (dislist != null && dislist.size() > 0) {
									for (int m = 0; m < dislist.size(); m++) {
										TCComponentBOMLine diebl = (TCComponentBOMLine) dislist.get(m);
										String diename = Util.getProperty(diebl, "bl_rev_object_name");
										String[] strVal = new String[3];
										int weldnum = 0;
										ArrayList weldlist = Util.getChildrenByBOMLine(diebl, "WeldPointRevision");
										if (weldlist != null && weldlist.size() > 0) {
											weldnum = weldlist.size();
										}
										// ���ݲ��������ж��Ƿ�ΪRH�У�����LH��
										if (lname.length() > 2
												&& lname.substring(lname.length() - 2).equals("LH")) {
											lhsum = lhsum + weldnum;
											strVal[2] = Integer.toString(weldnum);// LH
										} else {
											rhsum = rhsum + weldnum;
											strVal[1] = Integer.toString(weldnum);// RH
										}
										// ���ݵ㺸����������ΪR��M��ͷ��ͳ��
										if (diename.length() > 1 && (diename.substring(0, 1).equals("R")
												|| diename.substring(0, 1).equals("M"))) {
											rsw = rsw + weldnum;
											rswsum = rswsum + weldnum;
										} else {
											psw = psw + weldnum;
										}
										allsum = allsum + weldnum;
                                        //���������ֻ��һ����λ�����������
										if(flag2) {
											strVal[0] = lname + " " + statename;
										}else {
											strVal[0] = lname ;
										}
										
										dataList.add(strVal);

									}
								} else {
									String[] strVal = new String[3];
									 //���������ֻ��һ����λ�����������
									if(flag2) {
										strVal[0] = lname + " " + statename;
									}else {
										strVal[0] = lname ;
									}
									dataList.add(strVal);
								}
							}
						}
						if (flag) {
							String[] strVal = new String[3];
							strVal[0] = "�ܼ�";
							strVal[1] = Integer.toString(rhsum);
							strVal[2] = Integer.toString(lhsum);
							dataList.add(strVal);
						}
					}
				}
				map.put(linename, dataList);
				Integer[] strValue = new Integer[3];
				strValue[0] = psw + rsw;
				strValue[1] = psw;
				strValue[2] = rsw;
				statistics.put(linename, strValue);
			}
		}
		// �������ߴ���
		if (other != null && other.size() > 0) {
			ArrayList dataList = new ArrayList();
			int psw = 0;
			int rsw = 0;
			String linename = "����";//
			for (int j = 0; j < other.size(); j++) {
				TCComponentBOMLine linebl = (TCComponentBOMLine) other.get(j);
				int rhsum = 0;// RH �ܼ�����
				int lhsum = 0;// LH �ܼ�����
				boolean flag2 = Util.getIsMEProcStat(linebl);
				String lname = Util.getProperty(linebl.getItemRevision(), "b8_LineType") + Util.getProperty(linebl, "bl_rev_object_name"); // ʵ�ʲ�������
				ArrayList statelist = Util.getChildrenByBOMLine(linebl, "B8_BIWMEProcStatRevision");// ��װ��λ����
				if (statelist != null && statelist.size() > 0) {
					for (int k = 0; k < statelist.size(); k++) {
						TCComponentBOMLine statebl = (TCComponentBOMLine) statelist.get(k);
						String statename = Util.getProperty(statebl.parent(), "bl_rev_object_name"); // ��λ����
						ArrayList dislist = Util.getChildrenByBOMLine(statebl, "B8_BIWDiscreteOPRevision");// �㺸����
						if (dislist != null && dislist.size() > 0) {
							for (int m = 0; m < dislist.size(); m++) {
								TCComponentBOMLine diebl = (TCComponentBOMLine) dislist.get(m);
								String diename = Util.getProperty(diebl, "bl_rev_object_name");
								String[] strVal = new String[3];
								int weldnum = 0;
								ArrayList weldlist = Util.getChildrenByBOMLine(diebl, "WeldPointRevision");
								if (weldlist != null && weldlist.size() > 0) {
									weldnum = weldlist.size();
								}
								// ���ݲ��������ж��Ƿ�ΪRH�У�����LH��
								if (diename.length() > 2 && diename.substring(diename.length() - 2).equals("LH")) {
									lhsum = lhsum + weldnum;
									strVal[2] = Integer.toString(weldnum);// LH
								} else {
									rhsum = rhsum + weldnum;
									strVal[1] = Integer.toString(weldnum);// RH
								}
								// ���ݲ���������ΪR��M��ͷ��ͳ��
								if (diename.length() > 1 && (diename.substring(0, 1).equals("R")
										|| diename.substring(0, 1).equals("M"))) {
									rsw = rsw + weldnum;
									rswsum = rswsum + weldnum;
								} else {
									psw = psw + weldnum;
								}
								allsum = allsum + weldnum;

								 //���������ֻ��һ����λ�����������
								if(flag2) {
									strVal[0] = lname + " " + statename;
								}else {
									strVal[0] = lname ;
								}
								dataList.add(strVal);

							}
						} else {
							String[] strVal = new String[3];
							 //���������ֻ��һ����λ�����������
							if(flag2) {
								strVal[0] = lname + " " + statename;
							}else {
								strVal[0] = lname ;
							}
							dataList.add(strVal);
						}
					}
					String[] strVal = new String[3];
					strVal[0] = "�ܼ�";
					strVal[1] = Integer.toString(rhsum);
					strVal[2] = Integer.toString(lhsum);
					dataList.add(strVal);
				}
			}

			map.put(linename, dataList);
			Integer[] strValue = new Integer[3];
			strValue[0] = psw + rsw;
			strValue[1] = psw;
			strValue[2] = rsw;
			statistics.put(linename, strValue);
		}

		return map;
	}

	// ����BOP���㣬��ȡ���е������ߣ��������û�з��������ߣ���Ϊ����
	private ArrayList getAsahiLine(TCComponentBOMLine topbl) throws TCException {
		// TODO Auto-generated method stub
		ArrayList list = new ArrayList();
		AIFComponentContext[] chilrens = topbl.getChildren();
		for (AIFComponentContext chil : chilrens) {
			TCComponentBOMLine bl = (TCComponentBOMLine) chil.getComponent();
			// ���ݲ������Ƿ��в����ж��Ƿ�Ϊ���
			ArrayList xclist = Util.getChildrenByBOMLine(bl, "B8_BIWMEProcLineRevision");
			if (xclist != null && xclist.size() > 0) {
				list.add(bl);
			} else {
				other.add(bl);
			}
		}
		return list;
	}

}
