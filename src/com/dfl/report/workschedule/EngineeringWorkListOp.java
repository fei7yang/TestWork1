package com.dfl.report.workschedule;

import java.awt.Container;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dfl.report.ExcelReader.BoardInformation;
import com.dfl.report.ExcelReader.CoverInfomation;
import com.dfl.report.ExcelReader.baseinfoExcelReader;
import com.dfl.report.util.FileUtil;
import com.dfl.report.util.GenerateReportInfo;
import com.dfl.report.util.NewOutputDataToExcel;
import com.dfl.report.util.ReportUtils;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentFolderType;
import com.teamcenter.rac.kernel.TCComponentGroup;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentRole;
import com.teamcenter.rac.kernel.TCComponentTcFile;
import com.teamcenter.rac.kernel.TCComponentUser;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCProperty;
import com.teamcenter.rac.kernel.TCSession;

public class EngineeringWorkListOp {

	private AbstractAIFUIApplication app;
	SimpleDateFormat df = new SimpleDateFormat("yyyy��MM��");// �������ڸ�ʽ
	SimpleDateFormat df2 = new SimpleDateFormat("yyyy.MM");// �������ڸ�ʽ
	private String Edition;
	private String topfoldername;
	private Map<String, String> projVehMap;
	private String VehicleNo;
	private TCSession session;
	private List<TCComponentDataset> datasetList = new ArrayList<TCComponentDataset>();
	private List<TCComponentItemRevision> revlist = new ArrayList<TCComponentItemRevision>();
	private GenerateReportInfo info;
	private InputStream inputStream = null;

	public EngineeringWorkListOp(AbstractAIFUIApplication app, Object object, String edition, String topfoldername,
			GenerateReportInfo info, InputStream inputStream) throws TCException {
		// TODO Auto-generated constructor stub
		this.app = app;
		this.Edition = edition;
		this.topfoldername = topfoldername;
		session = (TCSession) app.getSession();
		this.info = info;
		this.inputStream = inputStream;
		initUI();
	}

	private void initUI() throws TCException {
		// TODO Auto-generated method stub
		InterfaceAIFComponent ift = app.getTargetComponent();
		TCComponentBOMLine topbl = (TCComponentBOMLine) ift;
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
		// �ļ�����
		String procName = "01.Ŀ¼";

		// ������ʾ���Ȳ����ִ�в���
		ReportViwePanel viewPanel = new ReportViwePanel("���ɱ���");
		viewPanel.setVisible(true);

		viewPanel.addInfomation("��ʼ�������...\n", 10, 100);

		viewPanel.addInfomation("", 20, 100);

		viewPanel.addInfomation("", 40, 100);

		// ����Ŀ¼
		String message2 = addNewReportContents(topbl, inputStream, info, procName, viewPanel);

		if (!message2.isEmpty()) {
			viewPanel.addInfomation(message2, 100, 100);
			return;
		}
		viewPanel.addInfomation("���������ɣ����ں�װ�������ն��󸽼��²鿴\n", 100, 100);
	}

	// ����Ŀ¼
	private String addNewReportContents(TCComponentBOMLine topbl, InputStream inputStream, GenerateReportInfo info,
			String procName, ReportViwePanel viewPanel) throws TCException {

		String error = "";

		// ���ݶ���BOP��ȡ���еĺ�װ��λ
		ArrayList dhlist = new ArrayList();
		getDiscretes(topbl, dhlist);

		ArrayList plist = new ArrayList();// ��ȡ�Ĺ�����������ݼ���
		ArrayList list = new ArrayList();// ������Ŀ¼������ݼ���
		ArrayList templist = new ArrayList();// ��ȡ��Ŀ¼������ݼ���

		List<BoardInformation> bzlist = new ArrayList<BoardInformation>();// ��ȡ�İ������ݼ���
		// ������
		// String username = app.getSession().getUserName();
		TCComponentUser user = session.getUser();
		String username = user.getUserName();
		TCComponentGroup group = session.getGroup();
		// ����
		String groupname = group.getLocalizedFullName();
		// ���п�
		String department = "";
		if (groupname != null
				&& (groupname.contains("ͬ�ڹ��̿�") || groupname.contains("simultaneous Engineering Section"))) {
			department = "H30";
		} else if (groupname != null
				&& (groupname.contains("��װ������") || groupname.contains("Body Assembly Engineering Section"))) {
			department = "VE2";
		} else {
			department = "VE2";
		}
		plist.add(username);
		// ����
		plist.add(df2.format(new Date()));
		// ����
		// ��Ϊ����BOP����ȡ����������Ϣ
		String objectname = Util.getProperty(topbl.getItemRevision(), "object_name");
		String factoryline = ReportUtils.getFactoryLineByBOP(objectname);
		String factory = "";
		String linebody = "";
		if (factoryline.length() > 3) {
			factory = factoryline.substring(0, 3);
			linebody = factoryline.substring(factoryline.length() - 1);
		}
		String car_type = VehicleNo + "-" + factory + "-NO" + linebody;

		plist.add(car_type);
		plist.add(Edition);
		plist.add(department);

		viewPanel.addInfomation("", 50, 100);

		ArrayList gxbh = new ArrayList();
		// ��ȡ��λ��Ϣ
		for (int i = 0; i < dhlist.size(); i++) {
			String[] str = new String[7];
			TCComponentBOMLine bomline = (TCComponentBOMLine) dhlist.get(i);
			TCComponentItemRevision blrev = bomline.getItemRevision();

			String stationcode = Util.getProperty(blrev, "b8_OPNo");
//			String[] assynos;
//			TCProperty p = blrev.getTCProperty("b8_AssyNo");
//			if (p != null) {
//				assynos = blrev.getTCProperty("b8_AssyNo").getStringValueArray(); // ������ ����
//			} else {
//				assynos = null;
//			}
//			if (assynos != null && assynos.length > 0) {
//				if (assynos[0].length() > 5) {
//					stationcode = "M" + assynos[0].trim().substring(0, 5);// ��λ���
//				} else {
//					stationcode = "M" + assynos[0].trim();
//				}
//
//			} else {
//				stationcode = "";// ��λ���
//			}

			str[0] = stationcode;// ������
			str[6] = Util.getProperty(bomline, "bl_rev_object_name");// ��λ����
			// str[1] = Util.getProperty(bomline.parent().getItemRevision(),
			// "b8_ChineseName");// ������������
			String chinesename = Util.getProperty(blrev, "b8_STName");
			str[1] = chinesename;
			String englishname = Util.getProperty(bomline.parent(), "bl_rev_object_name");// ����Ӣ������
			if (chinesename.contains(str[6])) {
				englishname = englishname + " " + str[6] + " ";
			}
			if (chinesename.contains("�Ҩu��")) {
				englishname = englishname.replace("RH", "").replace("LH", "") + "RH/LH";
			}
			if (chinesename.contains("��u��")) {
				englishname = englishname.replace("RH", "").replace("LH", "") + "LH/RH";
			}
			str[2] = englishname;
			str[3] = Edition;// ���
			str[4] = Util.getProperty(blrev, "b8_OpSheetNumber");// ҳ��b8_OpSheetNumber
			// ��Ϊ��ȡ���ɵĹ�����ҵ�����sheet��
//			String baseprocName = str[1].replace("#", "");
//			String baseprocName2 = str[1].replace("#", "");
//			InputStream inputS = null;
//			inputS = baseinfoExcelReader.getFileinbyreadExcel(blrev, "IMAN_specification", baseprocName);
//			if (inputS == null) {
//				inputS = baseinfoExcelReader.getFileinbyreadExcel(blrev, "IMAN_specification", baseprocName2);
//			}
//			if (inputS != null) {
//				XSSFWorkbook book = NewOutputDataToExcel.creatXSSFWorkbook(inputStream);
//				int sheetnum = book.getNumberOfSheets();
//				str[4] = Integer.toString(sheetnum);
//			} else {
//				str[4] = "0";
//			}
			str[5] = "";// ��ע

			// ���ҳ��Ϊ�գ�˵���ù�λδ������������
			if (str[4] != null && !str[4].isEmpty()) {
				if (!gxbh.contains(str[0])) {
					gxbh.add(str[0]);
				}
				templist.add(str);
			}
		}
		// ����λ��Ϣ�����������һ������Ҫ��Ϊһ����ʾ,ҳ��ϼ�
		for (int i = 0; i < gxbh.size(); i++) {
			String ProcessNumber = (String) gxbh.get(i);
			String[] value = new String[7];
			String pEnglishName = "";
			String pChineseName = "";
			int page = 0;
			String Edition = "";
			// ���ڱ��
			int num = 0;
			String tempName = "";
			String lastname = "";
			for (int j = 0; j < templist.size(); j++) {
				String[] str = (String[]) templist.get(j);
				if (str[0].equals(ProcessNumber)) {
					num++;
					if (num == 1) {
						pEnglishName = str[2];
						pChineseName = str[1];
						tempName = str[6];
					} else if (num == 2) {
						// tempName = tempName + "," + str[6];
						lastname = str[6];
					} else {
						// tempName = tempName + "~" + str[6];
						lastname = str[6];
					}
					if (str[4] != null && !str[4].isEmpty()) {
						page = page + Integer.parseInt(str[4]);
					}
					Edition = str[3];
				}
			}
			if (num == 1) {

			} else if (num == 2) {
				pChineseName = pChineseName.replace(tempName, "") + "(" + tempName + "," + lastname + ")";
				pEnglishName = pEnglishName.replace(tempName, "") + "(" + tempName + "," + lastname + ")";
			} else {
				pChineseName = pChineseName.replace(tempName, "") + "(" + tempName + "-" + lastname + ")";
				pEnglishName = pEnglishName.replace(tempName, "") + "(" + tempName + "-" + lastname + ")";
			}

			value[0] = Integer.toString(i + 1); // ���
			value[1] = ProcessNumber;
			value[2] = pChineseName;
			value[3] = pEnglishName;
			value[4] = Edition;
			value[5] = Integer.toString(page);
			value[6] = "";

			list.add(value);
		}
		viewPanel.addInfomation("��ʼд���ݣ������ĵȴ�...\n", 60, 100);
		// ��ȡ������Ϣ
		String basename = "222.������Ϣ";
		bzlist = getPartData(topbl, basename);

		String filename = procName;

		// ������·
		{
			Util.callByPass(session, true);
		}

		XSSFWorkbook book = NewOutputDataToExcel.creatXSSFWorkbook(inputStream, list, bzlist);
		NewOutputDataToExcel.writeDataToSheet(book, plist, list, bzlist);
		for (int i = 0; i < book.getNumberOfSheets(); i++) 
		{
			XSSFSheet sheet = book.getSheetAt(i);
			if(sheet.getSheetName().contains("��¼-24��������") || sheet.getSheetName().contains("��¼-���ж��ձ�"))
			{
				book.setPrintArea(i, 0, 114, 0, 51);
				PrintSetup printSetup = sheet.getPrintSetup();
				printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
				printSetup.setScale((short) 65);// �Զ������ţ��˴�100Ϊ������
				printSetup.setLandscape(true); // ��ӡ����true������false������(Ĭ��)
			}
			if(sheet.getSheetName().contains("��¼-255��������"))
			{
				book.setPrintArea(i, 0, 114, 0, 51);
				PrintSetup printSetup = sheet.getPrintSetup();
				printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);				
				printSetup.setScale((short) 71);// �Զ������ţ��˴�100Ϊ������
				printSetup.setFitHeight((short) 1);
				printSetup.setFitWidth((short) 1);
				printSetup.setLandscape(true); // ��ӡ����true������false������(Ĭ��)
			}
		}
	
		NewOutputDataToExcel.exportFile(book, filename);
		String fullFileName = FileUtil.getReportFileName(filename);
		viewPanel.addInfomation("", 80, 100);
		TCComponentDataset ds = Util.createDataset(session, filename, fullFileName, "MSExcelX", "excel");
		datasetList.add(ds);
		revlist.add(topbl.getItemRevision());
		try {
			TCComponentItem docunment = ReportUtils.afterGenerateReportAction(datasetList, revlist, info, procName, "",
					session);
			saveFileToFolder(docunment, topfoldername);

			// �ļ���ź��������
			TCProperty pdoc = docunment.getTCProperty("dfl9_vehiclePlant");
			if (pdoc != null) {
				if (pdoc != null) {
					pdoc.setStringValue(topfoldername);
					docunment.lock();
					docunment.save();
					docunment.unlock();
				}
			}

		} catch (TCException e) {
			e.printStackTrace();
			// EclipseUtils.info(Messages.FixtureOperation_15 + e.getMessage());
			return error;
		}
		// �ر���·
		{
			Util.callByPass(session, false);
		}
		return error;
	}

	// ��ȡ������Ϣ
	private List<BoardInformation> getPartData(TCComponentBOMLine topbl, String procName) {
		// TODO Auto-generated method stub
		List<BoardInformation> baseinfolist = new ArrayList<BoardInformation>();

		InputStream filein = null;
		try {
			filein = baseinfoExcelReader.getFileinbyreadExcel2(topbl.getItemRevision(), "IMAN_reference", procName);
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		baseinfolist = baseinfoExcelReader.readBZExcel(filein, "xlsx");

		return baseinfolist;
	}

	// ��ȡ��װ��λ
	private ArrayList getDiscretes(TCComponentBOMLine topbl, ArrayList dhlist) {
		// TODO Auto-generated method stub
		try {
			AIFComponentContext[] children = topbl.getChildren();
			for (AIFComponentContext child : children) {
				TCComponentBOMLine bl = (TCComponentBOMLine) child.getComponent();
				TCComponentItemRevision rev = bl.getItemRevision();
				if (rev.isTypeOf("B8_BIWMEProcStatRevision")) {
					dhlist.add(bl);
					continue;
				} else {
					getDiscretes(bl, dhlist);
				}
			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return dhlist;
	}

	/*
	 * �����ļ���Ŵ��������ļ�����home��,������00.�����ļ��У��ѷ����ĵ��ŵ�00.�����ļ�����
	 */
	private void saveFileToFolder(TCComponentItem document, String topfoldername) {
		// TODO Auto-generated method stub
		try {
			TCComponentUser user = session.getUser();
			TCComponentFolder homefolder = user.getHomeFolder();
			TCComponentFolder folder = null;
			TCComponentFolder childrenfolder = null;
			// ���ж��Ƿ��Ѿ������˸��ļ���
			AIFComponentContext[] icf = homefolder.getChildren();
			for (AIFComponentContext aif : icf) {
				TCComponent tcc = (TCComponent) aif.getComponent();
				String obejctname = Util.getProperty(tcc, "object_name");
				if (tcc.getType().equals("Folder") && obejctname.equals(topfoldername)) {
					folder = (TCComponentFolder) tcc;
					break;
				}
			}
			if (folder == null) {
				return;
			}
			TCComponentFolderType foldertype = (TCComponentFolderType) session.getTypeComponent("Folder");

			// ���ж��Ƿ��Ѿ�������01.Ŀ¼����¼�ļ���
			AIFComponentContext[] icf1 = folder.getChildren();
			for (AIFComponentContext aif : icf1) {
				TCComponent tcc = (TCComponent) aif.getComponent();
				String obejctname = Util.getProperty(tcc, "object_name");
				if (tcc.getType().equals("Folder") && obejctname.equals("01.Ŀ¼����¼")) {
					childrenfolder = (TCComponentFolder) tcc;
					break;
				}
			}
			if (childrenfolder == null) {
				childrenfolder = foldertype.create("01.Ŀ¼����¼", "", "Folder");
				folder.add("contents", childrenfolder);
				childrenfolder.add("contents", document);
			} else {
				// folder.add("contents", childrenfolder);
				AIFComponentContext[] icf3 = childrenfolder.getChildren();
				// ���Ƴ�
				if (icf3 != null) {
					for (AIFComponentContext aif : icf3) {
						TCComponent tcc = (TCComponent) aif.getComponent();
						childrenfolder.remove("contents", tcc);
					}
				}
				childrenfolder.add("contents", document);
			}

		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
