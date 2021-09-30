package com.dfl.report.workschedule;

import java.io.InputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import com.dfl.report.ExcelReader.BoardInformation;
import com.dfl.report.ExcelReader.WeldPointInfo;
import com.dfl.report.ExcelReader.baseinfoExcelReader;
import com.dfl.report.util.FileUtil;
import com.dfl.report.util.GenerateReportInfo;
import com.dfl.report.util.NewOutputDataToExcel;
import com.dfl.report.util.ReportUtils;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentBOMWindow;
import com.teamcenter.rac.kernel.TCComponentBOMWindowType;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentFolderType;
import com.teamcenter.rac.kernel.TCComponentGroup;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentItemType;
import com.teamcenter.rac.kernel.TCComponentQuery;
import com.teamcenter.rac.kernel.TCComponentQueryType;
import com.teamcenter.rac.kernel.TCComponentRole;
import com.teamcenter.rac.kernel.TCComponentTcFile;
import com.teamcenter.rac.kernel.TCComponentUser;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;

public class BasicInformationBZOp {

	private AbstractAIFUIApplication app;
	private ArrayList weld = new ArrayList();
	private List<WeldPointInfo> weldlist = new ArrayList<WeldPointInfo>();// ������Ϣ�������
	// private TCComponentItemType tcccomponentitemtype;
	private TCComponentBOMWindow bomWin;
	List<TCComponentDataset> datasetList = new ArrayList<TCComponentDataset>();
	List<TCComponentItemRevision> revlist = new ArrayList<TCComponentItemRevision>();
	private Map<String, String> projVehMap;// ��ȡ��ѡ��ʹ�����familycode�Ĺ�ϵ
	private String VehicleNo = "";// ���ʹ���
	private DecimalFormat format = new DecimalFormat("0.00");
	private DecimalFormat format1 = new DecimalFormat("0.0000");
	private boolean bzflag = false; // �����ж��Ƿ�Ϊ��һ�������������
	private GenerateReportInfo info;
	private InputStream inputStream;
	private Map<String, List<String>> MaterialMap;

	public BasicInformationBZOp(AbstractAIFUIApplication app, GenerateReportInfo info, InputStream inputStream, Map<String, List<String>> materialMap) throws TCException {
		// TODO Auto-generated constructor stub
		this.app = app;
		this.info = info;
		this.inputStream = inputStream;
		this.MaterialMap = materialMap;
		initUI();
	}

	private void initUI() throws TCException {
		InterfaceAIFComponent ifc = app.getTargetComponent();
		TCComponentBOMLine topbl = (TCComponentBOMLine) ifc;

		TCSession session = (TCSession) app.getSession();
		TCComponentUser user = session.getUser();

		// ��ȡ ��Ŀ-���� ��ѡ��
		projVehMap = ReportUtils.getDFL_Project_VehicleNo();
		String FamlilyCode = "";
		FamlilyCode = Util.getProperty(topbl, "bl_rev_project_ids");// ��������
		if (projVehMap.size() < 1) {
			VehicleNo = FamlilyCode;
		} else {
			VehicleNo = projVehMap.get(FamlilyCode);
			if (VehicleNo == null) {
				if (FamlilyCode != null) {
					VehicleNo = FamlilyCode;
				}
			}
		}		
		// �״����ɱ���
		// ������ʾ���Ȳ����ִ�в���
		ReportViwePanel viewPanel = new ReportViwePanel("���ɱ���");
		viewPanel.setVisible(true);

		viewPanel.addInfomation("���ڻ�ȡģ��...\n", 10, 100);

		// �ж��û�����ѡ�����Ƿ���дȨ��
//		boolean flag;
//		try {
//			flag = Util.hasWritePrivilege(session, topbl.getItemRevision());
//			if (!flag) {
//				viewPanel.addInfomation("�Ե�ǰBOP��װ�������հ汾û��дȨ�ޣ�...\n", 100, 100);
//				return;
//			}
//		} catch (TCException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}

		XSSFWorkbook book = null;
		String procName = "222.������Ϣ";

		if (info.getAction() == "create") { // �����
			viewPanel.addInfomation("�������������Ϣ-�����嵥��Ϣ��", 100, 100);
			return;
		} else {
			
			book = creatXSSFWorkbook(inputStream);
			
			viewPanel.addInfomation("��ʼ�������...\n", 20, 100);

			// ��հ����嵥
			clearBZDataToSheet(book);

			weldlist = baseinfoExcelReader.readWeldExcel(book, "xlsx");
			List bzList = new ArrayList();
			try {
				bzList = getSolutionPart(weldlist, topbl, session, viewPanel);
			} catch (TCException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			// д����sheet����
			writeBZDataToSheet(book, bzList);
		}
		// ������·
		{
			Util.callByPass(session, true);
		}
		viewPanel.addInfomation("��ʼд���ݣ������ĵȴ�...\n", 80, 100);
		String filename = procName;
		filename = filename.replaceAll("\\s*", "");
		NewOutputDataToExcel.exportFile(book, filename);

		String fullFileName = FileUtil.getReportFileName(filename);
		System.out.println(fullFileName);
		TCComponentDataset ds = Util.createDataset(session, filename, fullFileName, "MSExcelX", "excel");
		if (ds != null) {
			datasetList.add(ds);
		}
		try {
			revlist.add(topbl.getItemRevision());
		} catch (TCException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		viewPanel.addInfomation("", 90, 100);
		try {
			ReportUtils.afterGenerateReportAction(datasetList, revlist, info, procName, "", session);
		} catch (TCException e) {
			e.printStackTrace();
			// EclipseUtils.info(Messages.FixtureOperation_15 + e.getMessage());
			return;
		}
		// �ر���·
		{
			Util.callByPass(session, false);
		}
		viewPanel.addInfomation("���������ɣ����ں�װ�������ն��󸽼��²鿴\n", 100, 100);
		viewPanel.addInfomation("��ܰ��ʾ��������Ϣ-�����嵥��Ϣ���ɳɹ����´��������Ѱ����嵥��Ϣ���ǣ��������������\n", 100, 100);

	}

	/*
	 * ��հ�����Ϣ
	 */
	private void clearBZDataToSheet(XSSFWorkbook book) {
		// TODO Auto-generated method stub
		// ��������
		Font font = book.createFont();
		// font.setColor((short) 12);
		font.setFontName("����");
		font.setFontHeightInPoints((short) 10);
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
//		cellStyle2.setBorderBottom(XSSFCellStyle.BORDER_THIN);
//		cellStyle2.setBorderLeft(XSSFCellStyle.BORDER_NONE);
//		cellStyle2.setBorderRight(XSSFCellStyle.BORDER_THIN);
//		cellStyle2.setBorderTop(XSSFCellStyle.BORDER_THIN);
//		cellStyle2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����
//		cellStyle2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
//		cellStyle2.setFont(font);

		XSSFCellStyle cellStyle3 = null;
//		cellStyle3.setBorderBottom(XSSFCellStyle.BORDER_THIN);
//		cellStyle3.setBorderLeft(XSSFCellStyle.BORDER_THIN);
//		cellStyle3.setBorderRight(XSSFCellStyle.BORDER_NONE);
//		cellStyle3.setBorderTop(XSSFCellStyle.BORDER_THIN);
//		cellStyle3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����
//		cellStyle3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
//		cellStyle3.setFont(font);

		// ����-����sheet
		XSSFSheet sh = book.getSheetAt(1);
		int rownum = sh.getPhysicalNumberOfRows();
		for (int i = 0; i < rownum; i++) {
			setStringCellAndStyle(sh, "", 1 + i, 0, cellStyle1, 10); // ���
			setStringCellAndStyle(sh, "", 1 + i, 1, cellStyle1, Cell.CELL_TYPE_STRING);// �����
			setStringCellAndStyle(sh, "", 1 + i, 2, cellStyle1, Cell.CELL_TYPE_STRING);// ������
			setStringCellAndStyle(sh, "", 1 + i, 4, cellStyle2, Cell.CELL_TYPE_STRING);// �������
			setStringCellAndStyle(sh, "", 1 + i, 5, cellStyle1, Cell.CELL_TYPE_STRING);// ����
			setStringCellAndStyle(sh, "", 1 + i, 6, cellStyle3, Cell.CELL_TYPE_STRING);// ���
			setStringCellAndStyle(sh, "", 1 + i, 7, cellStyle2, Cell.CELL_TYPE_STRING);// ���λ
			setStringCellAndStyle(sh, "", 1 + i, 8, cellStyle3, 10);// ǿ��
			setStringCellAndStyle(sh, "", 1 + i, 9, cellStyle2, Cell.CELL_TYPE_STRING);// ǿ�ȵ�λ
			setStringCellAndStyle(sh, "", 1 + i, 10, cellStyle1, Cell.CELL_TYPE_STRING);// GA/GI
		}

		// �����嵥sheet
		XSSFSheet sh2 = book.getSheetAt(2);
		int rownum2 = sh.getPhysicalNumberOfRows();
		for (int i = 0; i < rownum2; i++) {
			setStringCellAndStyle(sh2, "", 1 + i, 0, cellStyle1, 10); // ���
			setStringCellAndStyle(sh2, "", 1 + i, 1, cellStyle1, Cell.CELL_TYPE_STRING);// ������
			setStringCellAndStyle(sh2, "", 1 + i, 2, cellStyle1, Cell.CELL_TYPE_STRING);// �����
			setStringCellAndStyle(sh2, "", 1 + i, 4, cellStyle2, Cell.CELL_TYPE_STRING);// �������
			setStringCellAndStyle(sh2, "", 1 + i, 5, cellStyle1, Cell.CELL_TYPE_STRING);// ����
			setStringCellAndStyle(sh2, "", 1 + i, 6, cellStyle3, Cell.CELL_TYPE_STRING);// ���
			setStringCellAndStyle(sh2, "", 1 + i, 7, cellStyle2, Cell.CELL_TYPE_STRING);// ���λ
			setStringCellAndStyle(sh2, "", 1 + i, 8, cellStyle3, 10);// ǿ��
			setStringCellAndStyle(sh2, "", 1 + i, 9, cellStyle2, Cell.CELL_TYPE_STRING);// ǿ�ȵ�λ
			setStringCellAndStyle(sh2, "", 1 + i, 10, cellStyle1, Cell.CELL_TYPE_STRING);// GA/GI
		}

	}

	/*
	 * д�����嵥����
	 */
	private void writeHDDataToSheet(XSSFWorkbook book, ArrayList hdlist) {
//		XSSFCellStyle style = book.createCellStyle();
//		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����
//		style.setAlignment(XSSFCellStyle.ALIGN_LEFT);
		XSSFSheet sheet = book.getSheetAt(3);
		int row = 1;
		for (int i = 0; i < hdlist.size(); i++) {
			String[] value = (String[]) hdlist.get(i);
			for (int j = 0; j < value.length; j++) {
				setStringCellAndStyle(sheet, value[j], row, j, null, Cell.CELL_TYPE_STRING);
			}
			row++;
		}
	}

	/*
	 * д���а���sheetҳ�Ĺ�����Ϣ
	 */
	private void writeBZDataToSheet(XSSFWorkbook book, List plist) {
		// TODO Auto-generated method stub
		// ��������
		Font font = book.createFont();
		// font.setColor((short) 12);
		font.setFontName("����");
		font.setFontHeightInPoints((short) 10);
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
//		cellStyle2.setBorderBottom(XSSFCellStyle.BORDER_THIN);
//		cellStyle2.setBorderLeft(XSSFCellStyle.BORDER_NONE);
//		cellStyle2.setBorderRight(XSSFCellStyle.BORDER_THIN);
//		cellStyle2.setBorderTop(XSSFCellStyle.BORDER_THIN);
//		cellStyle2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����
//		cellStyle2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
//		cellStyle2.setFont(font);

		XSSFCellStyle cellStyle3 = null;
//		cellStyle3.setBorderBottom(XSSFCellStyle.BORDER_THIN);
//		cellStyle3.setBorderLeft(XSSFCellStyle.BORDER_THIN);
//		cellStyle3.setBorderRight(XSSFCellStyle.BORDER_NONE);
//		cellStyle3.setBorderTop(XSSFCellStyle.BORDER_THIN);
//		cellStyle3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����
//		cellStyle3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
//		cellStyle3.setFont(font);

		// ����-����sheet
		XSSFSheet sh = book.getSheetAt(1);
		for (int i = 0; i < plist.size(); i++) {
			String[] value = (String[]) plist.get(i);
			setStringCellAndStyle(sh, value[0], 1 + i, 0, cellStyle1, 10); // ���
			setStringCellAndStyle(sh, value[2], 1 + i, 1, cellStyle1, Cell.CELL_TYPE_STRING);// �����
			setStringCellAndStyle(sh, value[1], 1 + i, 2, cellStyle1, Cell.CELL_TYPE_STRING);// ������
			setStringCellAndStyle(sh, value[3], 1 + i, 4, cellStyle2, Cell.CELL_TYPE_STRING);// �������
			setStringCellAndStyle(sh, value[4], 1 + i, 5, cellStyle1, Cell.CELL_TYPE_STRING);// ����
			if (Util.isNumber(value[5])) {
				setStringCellAndStyle(sh, value[5], 1 + i, 6, cellStyle3, 11);// ���
			} else {
				setStringCellAndStyle(sh, value[5], 1 + i, 6, cellStyle3, Cell.CELL_TYPE_STRING);// ���
			}
			setStringCellAndStyle(sh, value[6], 1 + i, 7, cellStyle2, Cell.CELL_TYPE_STRING);// ���λ
			setStringCellAndStyle(sh, value[7], 1 + i, 8, cellStyle3, 10);// ǿ��
			setStringCellAndStyle(sh, value[8], 1 + i, 9, cellStyle2, Cell.CELL_TYPE_STRING);// ǿ�ȵ�λ
			setStringCellAndStyle(sh, value[9], 1 + i, 10, cellStyle1, Cell.CELL_TYPE_STRING);// GA/GI
		}

		// �����嵥sheet
		XSSFSheet sh2 = book.getSheetAt(2);
		for (int i = 0; i < plist.size(); i++) {
			String[] value = (String[]) plist.get(i);
			setStringCellAndStyle(sh2, value[0], 1 + i, 0, cellStyle1, 10); // ���
			setStringCellAndStyle(sh2, value[1], 1 + i, 1, cellStyle1, Cell.CELL_TYPE_STRING);// ������
			setStringCellAndStyle(sh2, value[2], 1 + i, 2, cellStyle1, Cell.CELL_TYPE_STRING);// �����
			setStringCellAndStyle(sh2, value[3], 1 + i, 4, cellStyle2, Cell.CELL_TYPE_STRING);// �������
			setStringCellAndStyle(sh2, value[4], 1 + i, 5, cellStyle1, Cell.CELL_TYPE_STRING);// ����
			if (Util.isNumber(value[5])) {
				setStringCellAndStyle(sh2, value[5], 1 + i, 6, cellStyle3, 11);// ���
			} else {
				setStringCellAndStyle(sh2, value[5], 1 + i, 6, cellStyle3, Cell.CELL_TYPE_STRING);// ���
			}
			setStringCellAndStyle(sh2, value[6], 1 + i, 7, cellStyle2, Cell.CELL_TYPE_STRING);// ���λ
			setStringCellAndStyle(sh2, value[7], 1 + i, 8, cellStyle3, 10);// ǿ��
			setStringCellAndStyle(sh2, value[8], 1 + i, 9, cellStyle2, Cell.CELL_TYPE_STRING);// ǿ�ȵ�λ
			setStringCellAndStyle(sh2, value[9], 1 + i, 10, cellStyle1, Cell.CELL_TYPE_STRING);// GA/GI
		}

	}

	private List getSolutionPart(List<WeldPointInfo> weldlist2, TCComponentBOMLine topbl, TCSession session,
			ReportViwePanel viewPanel) throws TCException {
		// TODO Auto-generated method stub
		ArrayList bzqclist = new ArrayList();
		ArrayList partlist = new ArrayList();
		int rowNum = 0;// ���
		// ͨ������Ų��Ҷ�Ӧ���������
		if (weldlist != null) {
			// �԰������ݸ������������
			Comparator comparator = getComParatorBypartno();
			Collections.sort(weldlist, comparator);

			TCComponentItemRevision toprev = topbl.getItemRevision();
			TCComponent[] bbomlist = toprev.getRelatedComponents("IMAN_METarget");

			for (int i = 0; i < weldlist.size(); i++) {
				double sch = (i + 1.0) / weldlist.size();
				int s = (int) (sch * 60);
				if (s <= 60) {
					viewPanel.addInfomation("", 20 + s, 100);
				}
				WeldPointInfo weldinfo = weldlist.get(i);
				String partNo = weldinfo.getPartno();

				// ���������ȥ��
				if (!bzqclist.contains(partNo)) {
					bzqclist.add(partNo);
					String[] values = new String[10];
					values[0] = Integer.toString(rowNum + 1);// ���
					values[1] = "BZ" + String.format("%03d", rowNum + 1);// ������
					values[2] = partNo; // �����
					String partname = "";// �������
					System.out.println("������BBOM����Ϊ:" + bbomlist.toString());
					if (bbomlist != null && bbomlist.length > 0) {
						for (int j = 0; j < bbomlist.length; j++) {
							TCComponentItemRevision bbomrev = (TCComponentItemRevision) bbomlist[j];
							// TCSession session = (TCSession) app.getSession();
							TCComponentBOMLine root;
							// �����칤�չ滮���л�ȡ�򿪵�BBOM�ṹ
							String bbomID = Util.getProperty(bbomrev, "item_id");
							System.out.println("bbomID:" + bbomID);

							root = Util.getOpenBOMLine(bbomID);
							if (root == null) {
								TCComponentBOMWindowType bomWinType;
								bomWinType = (TCComponentBOMWindowType) session.getTypeComponent("BOMWindow");
								bomWin = bomWinType.create(null);
								root = bomWin.setWindowTopLine(null, bbomrev, null, null);
							}
							// ����ϵͳ��ѯ����ȡ��صİ��
							// ��������������˺�׺����Ҫ��ȡ���������ڲ�ѯ
							String querypartno = "";
							if (partNo.length() > 2) {
								String spcilchar = partNo.substring(partNo.length() - 2, partNo.length() - 1);
								if (spcilchar.equals("-")) {
									querypartno = partNo.substring(0, partNo.length() - 2);
								} else {
									querypartno = partNo;
								}
							}
							List tcclist = Util.callStructureSearch(root, "__DFL_Find_SolutionPart",
									new String[] { "PARTNO" }, new String[] { querypartno });
							if (tcclist != null && tcclist.size() > 0) {
								TCComponentBOMLine solbl = (TCComponentBOMLine) tcclist.get(0);
								// partname = Util.getProperty(solbl, "bl_rev_object_name");
								partname = Util.getProperty(solbl.getItemRevision(), "dfl9_CADObjectName");
								break; // �ҵ�һ����Ӧ��������ƾ��жϲ�ѯ
							}
							if (bomWin != null) {
								bomWin.close();
							}
						}
					}
					values[3] = partname;
					String Partmaterial = weldinfo.getPartmaterial();
					values[4] = Partmaterial; // ����
					values[5] = weldinfo.getPartthickness(); // ���
					if (values[5] != null && !values[5].isEmpty()) {
						values[6] = "mm";
					} else {
						values[6] = "";
					}

					// ���ݲ��ʻ�ȡǿ�Ⱥ�GA/GI������
					String Sheetstrength = "";// ǿ��
					String Gagi = "";// GA/GI
					// ����Ǻ񱡰壬�޷���ȡǿ�Ⱥ�GA/GI��
					boolean flag = getJudgingThickSheet(Partmaterial);
					if (!flag) {
						if (Partmaterial != null && !Partmaterial.isEmpty()) {
							String[] str = Partmaterial.split("-");
							if (str.length > 1) {
								String tempstr = str[1].trim();
								if (tempstr != null && !"".equals(tempstr)) {
									for (int K = 0; K < tempstr.length(); K++) {
										if (tempstr.charAt(K) >= 48 && tempstr.charAt(K) <= 57) {
											Sheetstrength += tempstr.charAt(K);
										}
									}
								}
							}
							if (!Sheetstrength.isEmpty() && Integer.parseInt(Sheetstrength) >= 440) {
								values[7] = Sheetstrength;
								values[8] = "mpa";
							} else {
								values[7] = "";
								values[8] = "";
							}
						}
//						if (Partmaterial != null && Partmaterial.length() > 4) {
//							String gagitem = Partmaterial.trim().substring(0, 4);
//							if (gagitem.equals("SP78") || gagitem.equals("SP79") || gagitem.equals("RP78")
//									|| gagitem.equals("RP79")) {
//								Gagi = "GA";
//							} else if (gagitem.equals("SP70") || gagitem.equals("SP71") || gagitem.equals("SP72")
//									|| gagitem.equals("SP73") || gagitem.equals("SP76") || gagitem.equals("RP70")
//									|| gagitem.equals("RP71") || gagitem.equals("RP72") || gagitem.equals("RP73")
//									|| gagitem.equals("RP76")) {
//								Gagi = "GI";
//							} else {
//								Gagi = "";
//							}
//						}
					} else {
						values[7] = "";
						values[8] = "";
					}
					//���ݲ��ʶ��ձ��ȡGA/GI����
					if(MaterialMap!=null)
					{
						for(Map.Entry<String, List<String>> entry: MaterialMap.entrySet())
						{
							String MaterialNo = entry.getKey();
							List<String> infolist = entry.getValue();
							if(Util.getIsEqueal(Partmaterial, MaterialNo))
							{
								if(!"-".equalsIgnoreCase(infolist.get(0)))
								{
									Gagi = infolist.get(0);
								}								
							}
							
						}
					}
					values[9] = Gagi;

					partlist.add(values);

					rowNum++;
				}
			}
		}else {
			viewPanel.addInfomation("", 50, 100);
		}

		return partlist;
	}

	// �ж��Ƿ�Ϊ�񱡰�
	private boolean getJudgingThickSheet(String partmaterial1) {
		// TODO Auto-generated method stub
		boolean flag = false;
		int count1 = 0;
		int count2 = 0;
		String str = "";
		if (partmaterial1 != null) {
			str = partmaterial1;
		}
		count1 = (str.length() - str.replace("SP", "").length()) / "SP".length();
		count2 = (str.length() - str.replace("RP", "").length()) / "RP".length();

		if (count1 + count2 > 1) {
			flag = true;
		}
		return flag;
	}

	private void getAllWeldPoint(TCSession session, TCComponentBOMLine topbl) {
		// TODO Auto-generated method stub

		System.out.println("���ڲ�������Ӧ����");
		// ��ȡBOP������BBOM����
		ArrayList qclist = new ArrayList();
		try {
			TCComponentItemRevision toprev = topbl.getItemRevision();
			System.out.println("���ڲ�������Ӧ����");
			TCComponent[] bbomlist = toprev.getRelatedComponents("IMAN_METarget");
			System.out.println("������BBOM����Ϊ:" + bbomlist.toString());
			if (bbomlist != null && bbomlist.length > 0) {
				// ����һ��Map�����ж��Ƿ�Ϊ��ͬ����������ظ���ѯ
				Map<String, String[]> partMap = new HashMap<String, String[]>();

				for (int i = 0; i < bbomlist.length; i++) {
					TCComponentItemRevision bbomrev = (TCComponentItemRevision) bbomlist[i];
					// TCSession session = (TCSession) app.getSession();
					TCComponentBOMLine root;
					// �����칤�չ滮���л�ȡ�򿪵�BBOM�ṹ
					String bbomID = Util.getProperty(bbomrev, "item_id");
					System.out.println("bbomID:" + bbomID);
					root = Util.getOpenBOMLine(bbomID);
					if (root == null) {
						TCComponentBOMWindowType bomWinType;
						bomWinType = (TCComponentBOMWindowType) session.getTypeComponent("BOMWindow");
						bomWin = bomWinType.create(null);
						root = bomWin.setWindowTopLine(null, bbomrev, null, null);
					}

					// ����BBOM��ѯ���еĺ���
					String weldtypename = Util.getObjectDisplayName(session, "WeldPoint");
					String[] propertys = new String[] { "bl_item_object_type", "bl_item_object_type" };
					String[] values = new String[] { weldtypename, weldtypename };

					ArrayList partList = Util.searchBOMLine(root, "OR", propertys, "==", values);
					System.out.println("�����ĺ��㣺" + partList.toString());
					if (partList != null && partList.size() > 0) {

						// ���ݺ������飬һ���Բ�ѯ���к�������İ��
//						TCComponentBOMLine[] partstr = new TCComponentBOMLine[partList.size()];
//						partList.toArray(partstr);
//						HashMap<TCComponentBOMLine,TCComponent[]> map = Util.getConnectedLines(session,partstr);

						for (int j = 0; j < partList.size(); j++) {
							String[] value = new String[17];
							String[] bzvalue = new String[5];
							TCComponentBOMLine bl = (TCComponentBOMLine) partList.get(j);
							TCComponentItemRevision rev = bl.getItemRevision();

							value[0] = Util.getProperty(bl, "bl_rev_object_name");// ������
							// ��ȡx,y,z����
							String xform = Util.getProperty(bl, "bl_plmxml_abs_xform");// ���Ա任����
							Double[] xyzArray = getXYZ(xform);
							Double x = xyzArray[0] * 1000;
							Double y = xyzArray[1] * 1000;
							Double z = xyzArray[2] * 1000;

							value[1] = format.format(new BigDecimal(x.toString()));// X����
							value[2] = format.format(new BigDecimal(y.toString()));// Y����
							value[3] = format.format(new BigDecimal(z.toString()));// Z����
							value[4] = "";// 1
							value[5] = "";// 2
							value[6] = "";// 3
							// ֻ��A��B��Ҫ�Ȳ���ʾ��������Ϊ��
							String important = Util.getProperty(rev, "b8_ImportantLevel");// ��Ҫ��b8_ImportantLevel
							if (important.equals("A") || important.equals("B")) {
								value[7] = important;// ��Ҫ��b8_ImportantLevel
							} else {
								value[7] = "";
							}
							// ��ȡ���1
							String cp1 = Util.getProperty(rev, "b8_ConnPart1");
							if (cp1 != null && !cp1.equals("")) {
								// ����ϵͳ��ѯ����ȡ��صİ��
								if (partMap.containsKey(cp1)) {
									String[] strvalue = partMap.get(cp1);
									value[8] = cp1;
									value[9] = strvalue[0];
									value[10] = strvalue[1];
								} else {
									// ����ϵͳ��ѯ����ȡ��صİ��
									String[] strvalue = getPropertysBypartNo(root, cp1);
									value[8] = cp1;
									value[9] = strvalue[0];
									value[10] = strvalue[1];
									partMap.put(cp1, strvalue);
								}
							}
							// ��ȡ���2
							String cp2 = Util.getProperty(rev, "b8_ConnPart2");
							if (cp2 != null && !cp2.equals("")) {
								if (partMap.containsKey(cp2)) {
									String[] strvalue = partMap.get(cp2);
									value[11] = cp2;
									value[12] = strvalue[0];
									value[13] = strvalue[1];
								} else {
									// ����ϵͳ��ѯ����ȡ��صİ��
									String[] strvalue = getPropertysBypartNo(root, cp2);
									value[11] = cp2;
									value[12] = strvalue[0];
									value[13] = strvalue[1];
									partMap.put(cp2, strvalue);
								}
							}
							// ��ȡ���3
							String cp3 = Util.getProperty(rev, "b8_ConnPart3");
							if (cp3 != null && !cp3.equals("")) {
								// ����ϵͳ��ѯ����ȡ��صİ��
								if (partMap.containsKey(cp3)) {
									String[] strvalue = partMap.get(cp3);
									value[14] = cp3;
									value[15] = strvalue[0];
									value[16] = strvalue[1];
								} else {
									// ����ϵͳ��ѯ����ȡ��صİ��
									String[] strvalue = getPropertysBypartNo(root, cp3);
									value[14] = cp3;
									value[15] = strvalue[0];
									value[16] = strvalue[1];
									partMap.put(cp3, strvalue);
								}
							}
							// ���ݺ�����ȥ��
//							if (!qclist.contains(value[0])) {
//								qclist.add(value[0]);
//								weld.add(value);
//							}
							weld.add(value);
						}
					}
					if (bomWin != null) {
						bomWin.close();
					}

				}
			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	// ���ò�ѯ��ȡ�������
	private String[] getPropertysBypartNo(TCComponentBOMLine root, String partno) throws TCException {
		String[] values = new String[2];
		// ����ϵͳ��ѯ����ȡ��صİ��
		List tcclist = Util.callStructureSearch(root, "__DFL_Find_SolutionPart", new String[] { "PARTNO" },
				new String[] { partno });
		if (tcclist != null && tcclist.size() > 0) {
			TCComponentBOMLine sol = (TCComponentBOMLine) tcclist.get(0);
			TCComponentItemRevision solrev3 = sol.getItemRevision();
			// values[0] = Util.getProperty(solrev3, "dfl9_part_no");// ����3
			String bh3 = Util.getProperty(solrev3, "dfl9PartThickness");// ���
			if (bh3 != null && !bh3.isEmpty()) {
				values[0] = format.format(new BigDecimal(bh3.toString()));
			} else {
				values[0] = bh3;
			}
			values[1] = Util.getProperty(solrev3, "dfl9PartMaterial");// ����
		}

		return values;
	}

	// ��ȡ��������꣨x,y,z��
	private Double[] getXYZ(String xform) {
		// TODO Auto-generated method stub
		Double[] values = new Double[] { 0.0, 0.0, 0.0 };
		String[] array = xform.split(" ");
		if (array != null && array.length == 16) {
			values[0] = Double.valueOf(array[12]);
			values[1] = Double.valueOf(array[13]);
			values[2] = Double.valueOf(array[14]);
		}
		return values;
	}

	// ����ģ�崴��Excel��ģ��
	public static XSSFWorkbook creatXSSFWorkbook(InputStream input) {
		XSSFWorkbook book = null;

		try {
			book = new XSSFWorkbook(input);
			XSSFSheet sheet1 = book.getSheetAt(3);
			//////////// ���÷�����ʾ�Ϸ�/�·�
			sheet1.setRowSumsBelow(false);
			sheet1.setRowSumsRight(false);
			sheet1.setRowSumsBelow(false);
			sheet1.setRowSumsRight(false);

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return book;

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

	private Comparator getComParatorBypartno() {
		// TODO Auto-generated method stub
		Comparator comparator = new Comparator() {

			public int compare(Object obj, Object obj1) {
				WeldPointInfo comp1 = (WeldPointInfo) obj;
				WeldPointInfo comp2 = (WeldPointInfo) obj1;

				String d1 = "";
				String d2 = "";
				if (comp1.getPartno() != null && !comp1.getPartno().isEmpty()) {
					d1 = comp1.getPartno();
				}
				if (comp2.getPartno() != null && !comp2.getPartno().isEmpty()) {
					d2 = comp2.getPartno();
				}
				return d1.compareTo(d2);
			}
		};

		return comparator;
	}
}
