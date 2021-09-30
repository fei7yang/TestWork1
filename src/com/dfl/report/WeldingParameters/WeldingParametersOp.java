package com.dfl.report.WeldingParameters;

import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.swt.widgets.Shell;

import com.dfl.report.common.TCComponentUtils;
import com.dfl.report.util.FileUtil;
import com.dfl.report.util.GenerateReportInfo;
import com.dfl.report.util.NewOutputDataToExcel;
import com.dfl.report.util.ReportUtils;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentUser;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCProperty;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.schemas.soa._2006_03.exceptions.ServiceException;
import com.teamcenter.services.rac.core._2008_06.DataManagement.CreateResponse;

public class WeldingParametersOp {

	private String Edition;
	private TCComponent savefolder;
	private TCSession session;
	private TCComponentBOMLine topbomline;
	SimpleDateFormat dateformat = new SimpleDateFormat("yyyy.MM");// �������ڸ�ʽ
	SimpleDateFormat dateformat2 = new SimpleDateFormat("yyyy��MM��");// �������ڸ�ʽ
	private List other = new ArrayList();
	List<TCComponentDataset> datasetList = new ArrayList<TCComponentDataset>();
	List<TCComponentItemRevision> revlist = new ArrayList<TCComponentItemRevision>();
	private InputStream inputStream;

	public WeldingParametersOp(TCComponentBOMLine topbomline, TCSession session, String edition, TCComponent savefolder, InputStream inputStream)
			throws TCException {
		// TODO Auto-generated constructor stub
		this.Edition = edition;
		this.savefolder = savefolder;
		this.session = session;
		this.topbomline = topbomline;
		this.inputStream = inputStream;
		initUI();
	}

	private void initUI() throws TCException {
		// TODO Auto-generated method stub

		// ������ʾ���Ȳ����ִ�в���
		ReportViwePanel viewPanel = new ReportViwePanel("���ɱ���");
		viewPanel.setVisible(true);

		viewPanel.addInfomation("��ʼ�������...\n", 10, 100);

	
		viewPanel.addInfomation("��ʼ��ȡ����...\n", 30, 100);
		String vehicle = Util.getProperty(topbomline, "bl_rev_project_ids");// ��������

		// ���ɱ������ǰ�Ķ���
//		GenerateReportInfo info = new GenerateReportInfo();
//		info.setExist(false);
//		info.setIsgoon(true);
//		info.setAction(""); //$NON-NLS-1$
//		info.setMeDocument(null);
//		info.setDFL9_process_type("H"); //$NON-NLS-1$
//		info.setDFL9_process_file_type("CS"); // $NON-NLS-1$
//		info.setmeDocumentName(procName);
//		info.setFlag(true);
//		try {
//			info = ReportUtils.beforeGenerateReportAction(topbomline.getItemRevision(), info);
//		} catch (TCException e) {
//			e.printStackTrace();
//			// EclipseUtils.info("Error : " + e.getMessage()); //$NON-NLS-1$
//			return;
//		}
//		System.out.println("The action is completed before the report operation is generated.");
//
//		if (!info.isIsgoon()) {
//			return;
//		}
//		// ֻ�������������Ǹ���
//		info.setAction("create");

		String BOPname = Util.getProperty(topbomline, "bl_rev_object_name");
		String filecode = "";
		String[] BOPnames = BOPname.split("_");
		String factory = "";
		if (BOPnames != null && BOPnames.length > 2) {
			vehicle = BOPnames[1];
			filecode = BOPnames[2];
			if (filecode.length() > 3) {
				factory = filecode.substring(0, 3);
			}
		}
		// ��ȡ���������Ϣ
		String procName = vehicle + "_" + Edition + "_PSW�������ܱ�"; // ����_�׶�_PSW�������ܱ�;
		// ����
		String date1 = dateformat2.format(new Date());
		TCComponentUser user = session.getUser();
		String username = user.getUserName();
		String[] cover = new String[6];
		cover[0] = "    ��    �ͣ�" + vehicle;
		cover[1] = "    ��    �Σ�" + Edition;
		cover[2] = "    �ļ���ţ�" + "��ͨ-" + filecode + "-CS";
		cover[3] = "    �������ڣ�" + date1;
		cover[4] = username;
		cover[5] = "    �������̣�" + factory + "������װ����";

		// ��ȡBOP�µ�������
		List obj = new ArrayList();
		List asshilist = getAsahiLine(topbomline);
		List factasshilist = new ArrayList();
		int totalpage = 1;
		viewPanel.addInfomation("", 40, 100);
		if (asshilist != null && asshilist.size() > 0) {		
			int count = 0;
			for (int i = 0; i < asshilist.size(); i++) {
				TCComponentBOMLine xuproc = (TCComponentBOMLine) asshilist.get(i);
				List gunlist = new ArrayList();
				List humangun = getHumanGunList(xuproc);
				//boolean Isweldpoint = getWeldPointList(xuproc);
				if (humangun != null && humangun.size() > 0) {
					factasshilist.add(xuproc);
					for (int j = 0; j < humangun.size(); j++) {
						String[] strValue = new String[17];
						TCComponentBOMLine gun = (TCComponentBOMLine) humangun.get(j);
						TCComponentItemRevision gunrev = gun.getItemRevision();

//						strValue[0] = Util.getProperty(gunrev, "b8_AdapterModel");// ��ѹ���ͺ�
//						strValue[1] = Util.getProperty(gunrev, "b8_Model");// ��ǹ�ͺ�
						// ��ȡ��Ӧ��ʵ�ʲ���
						TCComponentBOMLine procbl = gun.parent().parent().parent();
						boolean flag = Util.getIsMEProcStat(procbl);
						if (flag) {
							strValue[2] = Util.getProperty(procbl.getItemRevision(), "b8_ChineseName")
									+ Util.getProperty(gun.parent().parent(), "bl_rev_object_name");// ʹ�ù�λ
						} else {
							strValue[2] = Util.getProperty(procbl.getItemRevision(), "b8_ChineseName");// ʹ�ù�λ
						}
						TCComponentItemRevision diearrev = gun.parent().getItemRevision();
						String diearname = Util.getProperty(diearrev, "object_name");
						// ��ѹ����źͺ�ǹ��ţ���Ϊ�ӵ㺸����������ȡֵ
						String[] nameArr = diearname.split("\\\\");
						String TransformerNumber = "";
						String Guncode = "";
						TransformerNumber = nameArr[0];
						if (nameArr.length > 1) {
							Guncode = nameArr[1];
						}
						strValue[0] = TransformerNumber;
						strValue[1] = Guncode;
						strValue[3] = "��";// ����
						strValue[4] = Util.getProperty(gun.getItemRevision(), "b8_ElectrodeVol");// ��ǯ�ѹ��
						strValue[5] = Util.getProperty(diearrev, "b8_WeldForce");// ��ѹ��
						strValue[6] = "15";// Ԥѹʱ��
						strValue[7] = Util.getProperty(diearrev, "b8_RiseTime");// ����ʱ��
						strValue[8] = Util.getProperty(diearrev, "b8_CurrentTime1");// ��һͨ��ʱ��
						strValue[9] = Util.getProperty(diearrev, "b8_Current1");// ��һͨ�����
						strValue[10] = Util.getProperty(diearrev, "b8_Cool1");// ��ȴʱ��һ
						strValue[11] = Util.getProperty(diearrev, "b8_CurrentTime2");// �ڶ�ͨ��ʱ��
						strValue[12] = Util.getProperty(diearrev, "b8_Current2");// �ڶ�ͨ�����
						strValue[13] = Util.getProperty(diearrev, "b8_Cool2");// ��ȴʱ���
						strValue[14] = Util.getProperty(diearrev, "b8_CurrentTime3");// ����ͨ��ʱ��
						strValue[15] = Util.getProperty(diearrev, "b8_Current3");// ����ͨ�����
						strValue[16] = Util.getProperty(diearrev, "b8_KeepTime");// ����
						gunlist.add(strValue);
					}
					
					count++;
					
					// ���ݹ�λ����
					Comparator comparator2 = getComParatorBygwname();
					Collections.sort(gunlist, comparator2);
					
					obj.add(gunlist);
				}
				
			}
		}
		totalpage = factasshilist.size();
		// ������Ϣ
		String[] common = new String[5];
		common[0] = username;
		common[1] = dateformat.format(new Date());
		common[2] = Edition;
		common[3] = Integer.toString(totalpage);
		String fatory = " һ���� NO1";
		if (filecode != null && filecode.length() > 4) {
			String prefactory = filecode.substring(0, 2);
			String math = filecode.substring(2, 3);
			// ��������ת��
			String upermath = getUpperMath(math);
			String after = filecode.substring(filecode.length() - 1);
			fatory = prefactory + " " + upermath + "���� " + "NO" + after;

		}
		common[4] = fatory;
		viewPanel.addInfomation("", 50, 100);
		// �������������������ؿ�ģ��
		XSSFWorkbook book = creatEngineeringXSSFWorkbook(inputStream, factasshilist);

		// ��ʼд������
		writeDataToSheet(book, cover, obj, vehicle, common);

		viewPanel.addInfomation("��ʼд���ݣ������ĵȴ�...\n", 60, 100);
		String filename = Util.formatString(procName);

		NewOutputDataToExcel.exportFile(book, filename);

		String fullFileName = FileUtil.getReportFileName(filename);
		System.out.println(fullFileName);
		TCComponentDataset ds = Util.createDataset(session, filename, fullFileName, "MSExcelX", "excel");
//		if (ds != null) {
//			datasetList.add(ds);
//		}
//		revlist.add(topbomline.getItemRevision());
		try {
			TCComponentItem docunment = AddDocumentItem(ds, procName);

			viewPanel.addInfomation("", 80, 100);

			savefolder.add("contents", docunment);
			// �ļ���ź��������
			TCProperty pdoc = docunment.getTCProperty("dfl9_vehiclePlant");
			if (pdoc != null) {
				pdoc.setStringValue(fatory.replace(" ", ""));
				docunment.lock();
				docunment.save();
				docunment.unlock();
			}
		} catch (TCException e) {
			e.printStackTrace();
			// EclipseUtils.info(Messages.FixtureOperation_15 + e.getMessage());
			return;
		}
		viewPanel.addInfomation("���������ɣ�������ѡ�ļ����²鿴����...\n", 100, 100);
	}

	private boolean getWeldPointList(TCComponentBOMLine bl) throws TCException {
		// TODO Auto-generated method stub
		List list = new ArrayList();
		String weldtypename = Util.getObjectDisplayName(session, "WeldPoint");
		String[] propertys = new String[] { "bl_item_object_type", "bl_item_object_type" };
		String[] values = new String[] { weldtypename, weldtypename };

		ArrayList weldlist = Util.searchBOMLine(bl, "OR", propertys, "==", values);
		if (weldlist != null && weldlist.size() > 0) {
			for (int i = 0; i < weldlist.size(); i++) {
				// ���ݵ㺸���������ж��Ƿ�Ϊ�˹���ǹ
				TCComponentBOMLine weld = (TCComponentBOMLine) weldlist.get(i);
				String direaname = Util.getProperty(weld.parent(), "bl_rev_object_name");
				if (weld.parent().getItemRevision().isTypeOf("B8_BIWDiscreteOPRevision")
						&& !direaname.substring(0, 1).equals("R")) {
					return true;
				}
			}
		}
		return false;
	}

	private TCComponentItem AddDocumentItem(TCComponentDataset ds, String procName) throws TCException {
		// TODO Auto-generated method stub
		// ������excel�ļ�����ΪMSExcelX���ݼ���ʹ�ù���ϵ���ص�һ���½�MEDocument�汾
		TCComponentItem docuitem = null;
		Map<String, Object> itemMap = new HashMap<String, Object>();
		Map<String, Object> itemRevisionMap = new HashMap<String, Object>();
		Map<String, Object> itemRevMasterFormMap = new HashMap<String, Object>();
		itemMap.put("item_id", ""); //$NON-NLS-1$ //$NON-NLS-2$
		itemMap.put("object_name", procName); //$NON-NLS-1$
		itemMap.put("object_desc", ""); //$NON-NLS-1$
		itemMap.put("object_type", "DFL9MEDocument"); //$NON-NLS-1$
		itemRevisionMap.put("object_type", "DFL9MEDocumentRevision"); //$NON-NLS-1$
		itemRevisionMap.put("object_name", procName); //$NON-NLS-1$
		itemRevisionMap.put("dfl9_process_type", "H"); //$NON-NLS-1$
		itemRevisionMap.put("dfl9_process_file_type", "CS"); //$NON-NLS-1$
		itemRevMasterFormMap.put("object_type", "DFL9MEDocumentRevisionMaster"); //$NON-NLS-1$

		try {
			CreateResponse respose = TCComponentUtils.create(itemMap, itemRevisionMap, itemRevMasterFormMap);
			int num = respose.serviceData.sizeOfCreatedObjects();
			if (num > 0) {
				for (int i = 0; i < num; i++) {
					TCComponent comp = respose.serviceData.getCreatedObject(i);
					if (comp instanceof TCComponentItemRevision) {
						TCComponentItemRevision docuitemrev = (TCComponentItemRevision) comp;
						docuitemrev.add("IMAN_specification", ds);
						docuitem = docuitemrev.getItem();
					}

				}
			}
		} catch (ServiceException e) {
			e.printStackTrace();
			// throw new TCException("Create " + ReportUtils.DFL9MEDocument + " Fail : "
			// +e.getMessage()); //$NON-NLS-1$ //$NON-NLS-2$
		}
		return docuitem;
	}

	private String getUpperMath(String math) {
		// TODO Auto-generated method stub
		String value = "";
		String[] Uppernumbers = { "һ", "��", "��", "��", "��", "��", "��", "��", "��" };
		if (Util.isNumber(math)) {
			int num = Integer.parseInt(math);
			value = Uppernumbers[num - 1];
		}
		return value;
	}

	private void writeDataToSheet(XSSFWorkbook book, String[] cover, List obj, String vehicle, String[] common) {
		// TODO Auto-generated method stub
		// ��д�������Ϣ
		// ��������
		Font font = book.createFont();
		font.setFontName("������");
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);// �Ӵ�
		font.setFontHeightInPoints((short) 16);
		// ����һ����ʽ
		XSSFCellStyle cellStyle = book.createCellStyle();
		cellStyle.setBorderBottom(XSSFCellStyle.BORDER_NONE); // �±߿�
		cellStyle.setBorderLeft(XSSFCellStyle.BORDER_NONE);// ��߿�
		cellStyle.setBorderTop(XSSFCellStyle.BORDER_NONE);// �ϱ߿�
		cellStyle.setBorderRight(XSSFCellStyle.BORDER_NONE);// �ұ߿�
		cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����
		cellStyle.setAlignment(XSSFCellStyle.ALIGN_LEFT);// �����
		cellStyle.setFont(font);

		Font font2 = book.createFont();
		font2.setFontName("������");
		// font2.setBoldweight(Font.BOLDWEIGHT_BOLD);// �Ӵ�
		font2.setFontHeightInPoints((short) 16);
		// ����һ����ʽ
		XSSFCellStyle cellStyle2 = book.createCellStyle();
		cellStyle2.setBorderBottom(XSSFCellStyle.BORDER_NONE); // �±߿�
		cellStyle2.setBorderLeft(XSSFCellStyle.BORDER_NONE);// ��߿�
		cellStyle2.setBorderTop(XSSFCellStyle.BORDER_NONE);// �ϱ߿�
		cellStyle2.setBorderRight(XSSFCellStyle.BORDER_NONE);// �ұ߿�
		cellStyle2.setVerticalAlignment(XSSFCellStyle.VERTICAL_TOP);// ���ݴ�ֱ����
		// cellStyle2.setAlignment(XSSFCellStyle.ALIGN_LEFT);// �����
		cellStyle2.setFont(font2);

		Font font3 = book.createFont();
		font3.setFontName("����");
		// font2.setBoldweight(Font.BOLDWEIGHT_BOLD);// �Ӵ�
		font3.setFontHeightInPoints((short) 10);
		// ����һ����ʽ
		XSSFCellStyle cellStyle3 = book.createCellStyle();
		cellStyle3.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
		cellStyle3.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
		cellStyle3.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
		cellStyle3.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
		cellStyle3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����
		cellStyle3.setAlignment(XSSFCellStyle.ALIGN_CENTER);// ���ݴ�ֱ����
		cellStyle3.setFont(font3);

		// ����һ����ʽ
		XSSFCellStyle cellStyle4 = book.createCellStyle();
		cellStyle4.setFillForegroundColor(IndexedColors.RED.getIndex());
		cellStyle4.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellStyle4.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
		cellStyle4.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
		cellStyle4.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
		cellStyle4.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
		cellStyle4.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����
		cellStyle4.setAlignment(XSSFCellStyle.ALIGN_CENTER);// ���ݴ�ֱ����
		cellStyle4.setFont(font3);

		// ��������
		Font font4 = book.createFont();
		font4.setColor((short) 12);
		font4.setFontName("����");
		font4.setFontHeightInPoints((short) 16);
		// ����һ����ʽ
		XSSFCellStyle cellStyle5 = book.createCellStyle();
		cellStyle5.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle5.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle5.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle5.setBorderTop(XSSFCellStyle.BORDER_THIN);
		cellStyle5.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����
		cellStyle5.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		cellStyle5.setFont(font4);

		// ��������
		Font font5 = book.createFont();
		font5.setColor((short) 12);
		font5.setFontName("MS PGothic");
		font5.setBoldweight(Font.BOLDWEIGHT_BOLD); // ����Ӵ�
		font5.setFontHeightInPoints((short) 12);
		// ����һ����ʽ
		XSSFCellStyle cellStyle6 = book.createCellStyle();
		cellStyle6.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle6.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle6.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle6.setBorderTop(XSSFCellStyle.BORDER_THIN);
		cellStyle6.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����
		cellStyle6.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		cellStyle6.setFont(font5);

		// ����һ����ʽ
		XSSFCellStyle cellStyle61 = book.createCellStyle();
		cellStyle61.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle61.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		//cellStyle61.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle61.setBorderTop(XSSFCellStyle.BORDER_THIN);
		cellStyle61.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����
		cellStyle61.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		cellStyle61.setFont(font5);

		// ����һ����ʽ
		XSSFCellStyle cellStyle62 = book.createCellStyle();
		cellStyle62.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		//cellStyle62.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle62.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle62.setBorderTop(XSSFCellStyle.BORDER_THIN);
		cellStyle62.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����
		cellStyle62.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		cellStyle62.setFont(font5);

		// ��������
		Font font6 = book.createFont();
		font6.setColor((short) 12);
		font6.setFontName("����");
		font6.setFontHeightInPoints((short) 28);
		font6.setUnderline(Font.U_SINGLE);// �����»���
		// ����һ����ʽ
		XSSFCellStyle cellStyle7 = book.createCellStyle();
		cellStyle7.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle7.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle7.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle7.setBorderTop(XSSFCellStyle.BORDER_THIN);
		cellStyle7.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����
		cellStyle7.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		cellStyle7.setFont(font6);

		Font font7 = book.createFont();
		font7.setFontName("����");
		// font2.setBoldweight(Font.BOLDWEIGHT_BOLD);// �Ӵ�
		font7.setFontHeightInPoints((short) 8);
		// ����һ����ʽ
		XSSFCellStyle cellStyle8 = book.createCellStyle();
		cellStyle8.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
		cellStyle8.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
		cellStyle8.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
		cellStyle8.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
		cellStyle8.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����
		cellStyle8.setAlignment(XSSFCellStyle.ALIGN_CENTER);// ���ݴ�ֱ����
		cellStyle8.setFont(font7);

		XSSFSheet sheet = book.getSheetAt(0);

		setStringCellAndStyle(sheet, cover[5], 5, 3, cellStyle, Cell.CELL_TYPE_STRING);
		setStringCellAndStyle(sheet, cover[0], 6, 3, cellStyle, Cell.CELL_TYPE_STRING);
		setStringCellAndStyle(sheet, cover[1], 7, 3, cellStyle, Cell.CELL_TYPE_STRING);
		setStringCellAndStyle(sheet, cover[2], 8, 3, cellStyle, Cell.CELL_TYPE_STRING);
		setStringCellAndStyle(sheet, cover[3], 10, 3, cellStyle, Cell.CELL_TYPE_STRING);
		setStringCellAndStyle(sheet, cover[4], 12, 1, cellStyle2, Cell.CELL_TYPE_STRING);
        
		
		if (obj != null && obj.size() > 0) {
			for (int i = 0; i < obj.size(); i++) {							
					XSSFSheet sh = book.getSheetAt(1 + i);
					List data = (List) obj.get(i);
					setStringCellAndStyle(sh, vehicle, 6, 25, cellStyle3, Cell.CELL_TYPE_STRING);

					setStringCellAndStyle(sh, common[0], 2, 6, cellStyle5, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sh, common[1], 2, 30, cellStyle5, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sh, common[2], 50, 110, cellStyle6, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sh, Integer.toString(i + 1), 52, 110, cellStyle61, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sh, common[3], 52, 116, cellStyle62, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sh, common[4] + "    " + "PSW���Ӳ�����", 0, 37, cellStyle7, Cell.CELL_TYPE_STRING);
					if (data != null && data.size() > 0) {
						// ��ʼ��
						String beginstr = "";
						int beginrow = 9;// ��ʼ��
						int endrow = 0;// ��ֹ��
						int n = 0;// �����
						// �ϲ���Ԫ��
						CellRangeAddress region1;
						for (int j = 0; j < data.size(); j++) {
							String[] values = (String[]) data.get(j);
							setStringCellAndStyle(sh, Integer.toString(j + 1), 9 + j, 1, cellStyle3, 10);
							setStringCellAndStyle(sh, values[0], 9 + j, 4, cellStyle3, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[1], 9 + j, 8, cellStyle8, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[2], 9 + j, 15, cellStyle3, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[3], 9 + j, 25, cellStyle3, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[4], 9 + j, 46, cellStyle3, 11);
							// ����ѹ�������ڡ���ǯ�ѹ��������ɫ���
							double hs = getDoubleByString(values[4]);
							double js = getDoubleByString(values[5]);
							if (Util.isNumber(values[4]) && Util.isNumber(values[5])) {
								if (js > hs) {
									setStringCellAndStyle(sh, values[5], 9 + j, 50, cellStyle4, 11);
								} else {
									setStringCellAndStyle(sh, values[5], 9 + j, 50, cellStyle3, 11);
								}
							} else {
								setStringCellAndStyle(sh, values[5], 9 + j, 50, cellStyle3, 11);
							}

							setStringCellAndStyle(sh, values[6], 9 + j, 56, cellStyle3, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[7], 9 + j, 62, cellStyle3, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[8], 9 + j, 68, cellStyle3, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[9], 9 + j, 74, cellStyle3, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[10], 9 + j, 80, cellStyle3, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[11], 9 + j, 86, cellStyle3, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[12], 9 + j, 92, cellStyle3, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[13], 9 + j, 98, cellStyle3, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[14], 9 + j, 104, cellStyle3, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[15], 9 + j, 110, cellStyle3, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[16], 9 + j, 116, cellStyle3, Cell.CELL_TYPE_STRING);

							if (j == 0) {
								beginstr = values[2];
							} else {
								/**
								 * ******************************** ��λ�кϲ�
								 */
								if (!values[2].equals(beginstr)) {
									endrow = beginrow + n;
									region1 = new CellRangeAddress(beginrow, endrow, (short) 15, (short) 24);
									sh.addMergedRegion(region1);
									beginstr = values[2].toString();
									beginrow = endrow + 1;
									n = 0;
								} else {
									n++;
								}
								if (j == data.size() - 1) {
									endrow = beginrow + n;
									region1 = new CellRangeAddress(beginrow, endrow, (short) 15, (short) 24);
									sh.addMergedRegion(region1);
								}
							}
							if(data.size() == 1) {
								region1 = new CellRangeAddress(9, 9, (short) 15, (short) 24);
								sh.addMergedRegion(region1);
							}
						}
					}
				
				
			}
		}

		for (int i = 1; i < book.getNumberOfSheets(); i++) {
			XSSFSheet sht = book.getSheetAt(i);
			book.setPrintArea(i, 0, 123, 0, 54);
			PrintSetup printSetup = sht.getPrintSetup();
			printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
			printSetup.setScale((short) 62);// �Զ������ţ��˴�100Ϊ������
			printSetup.setLandscape(true); // ��ӡ����true������false������(Ĭ��)
		}
	}

	private Comparator getComParatorBygwname() {
		// TODO Auto-generated method stub
		Comparator comparator = new Comparator() {

			public int compare(Object obj, Object obj1) {
				// System.setProperty("java.util.Arrays.useLegacyMergeSort", "true");
				String[] comp1 = (String[]) obj;
				String[] comp2 = (String[]) obj1;

				String d1 = "";
				String d2 = "";
				if (obj != null && comp1[2] != null && !comp1[2].isEmpty()) {
					d1 = comp1[2].toString();
				}
				if (obj1 != null && comp2[2] != null && !comp2[2].isEmpty()) {
					d2 = comp2[2];
				}
				return d1.compareTo(d2);
			}
		};

		return comparator;
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
				if(Util.isNumber(value)) {
					cell.setCellType(Cell.CELL_TYPE_NUMERIC);
					cell.setCellValue(Integer.parseInt(value));
				}else {
					cell.setCellType(Cell.CELL_TYPE_STRING);
					cell.setCellValue(value);
				}
				
			} else if (celltype == 11) {
				if(Util.isNumber(value)) {
					cell.setCellType(Cell.CELL_TYPE_NUMERIC);
					cell.setCellValue(Double.parseDouble(value));
				}else {
					cell.setCellType(Cell.CELL_TYPE_STRING);
					cell.setCellValue(value);
				}
				
			} else {
				cell.setCellType(Cell.CELL_TYPE_STRING);
				cell.setCellValue(value);
			}
		}

		//cell.setCellStyle(Style);

	}

	private XSSFWorkbook creatEngineeringXSSFWorkbook(InputStream inputStream, List asshilist) {
		// TODO Auto-generated method stub
		XSSFWorkbook book = null;
		try {
			book = new XSSFWorkbook(inputStream);
			if (asshilist != null && asshilist.size() > 0) {
				for (int i = 0; i < asshilist.size(); i++) {
					TCComponentBOMLine xuproc = (TCComponentBOMLine) asshilist.get(i);
					String sheetname = Util.getProperty(xuproc, "bl_rev_object_name");
					String[] str = sheetname.split("[-_ ]");
					if(str!=null && str.length>1) {
						sheetname = str[1];
					}
					sheetname = String.format("%02d", i + 1) + sheetname;
					if (i == 0) {
						book.setSheetName(1, sheetname);
					} else {
						XSSFSheet newsheet = book.cloneSheet(1);
						int sheetat = book.getSheetIndex(newsheet);
						book.setSheetName(sheetat, sheetname);
						// book.setSheetOrder(newsheet.getSheetName(), index);
					}
				}
			}

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return book;
	}

	// ���������߻�ȡ���µ��˹���ǹ
	private List getHumanGunList(TCComponentBOMLine bl) throws TCException {
		List list = new ArrayList();
		String guntypename = Util.getObjectDisplayName(session, "B8_BIWGun");
		String[] propertys = new String[] { "bl_item_object_type", "bl_item_object_type" };
		String[] values = new String[] { guntypename, guntypename };

		ArrayList gunlist = Util.searchBOMLine(bl, "OR", propertys, "==", values);
		if (gunlist != null && gunlist.size() > 0) {
			for (int i = 0; i < gunlist.size(); i++) {
				// ���ݵ㺸���������ж��Ƿ�Ϊ�˹���ǹ
				TCComponentBOMLine gun = (TCComponentBOMLine) gunlist.get(i);
				String direaname = Util.getProperty(gun.parent(), "bl_rev_object_name");
				if (gun.parent().getItemRevision().isTypeOf("B8_BIWDiscreteOPRevision")
						&& !direaname.substring(0, 1).equals("R")) {
					list.add(gun);
				}
			}
		}
		return list;
	}

	// ����BOP���㣬��ȡ���е������ߣ��������û�з��������ߣ���Ϊ����
	private List getAsahiLine(TCComponentBOMLine topbl) throws TCException {
		// TODO Auto-generated method stub
		List list = new ArrayList();
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

	private double getDoubleByString(String str) {
		double dd = 0;
		if (str == null || !Util.isNumber(str)) {
			dd = 0;
		} else {
			dd = Double.parseDouble(str);
		}

		return dd;
	}
}
