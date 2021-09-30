package com.dfl.report.workschedule;

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
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dfl.report.ExcelReader.baseinfoExcelReader;
import com.dfl.report.common.EclipseUtils;
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

public class EngineeringWorkListCoverOp {

	private AbstractAIFUIApplication app;
	SimpleDateFormat df = new SimpleDateFormat("yyyy��MM��");// �������ڸ�ʽ
	SimpleDateFormat df2 = new SimpleDateFormat("yyyy.MM");// �������ڸ�ʽ
	private Map<String, String> projVehMap;
	private String VehicleNo = "";// ���ʹ���
	private String Edition = "";// ���
	List<TCComponentDataset> datasetList = new ArrayList<TCComponentDataset>();
	List<TCComponentItemRevision> revlist = new ArrayList<TCComponentItemRevision>();
	private TCSession session;
	private InputStream inputStream = null;
	private GenerateReportInfo info;

	public EngineeringWorkListCoverOp(AbstractAIFUIApplication app, Object object, String edition2, GenerateReportInfo info, InputStream inputStream) throws TCException {
		// TODO Auto-generated constructor stub
		this.app = app;
		this.Edition = edition2;
		session = (TCSession) app.getSession();
		this.inputStream = inputStream;
		this.info = info ;
		initUI();
	}

	private void initUI() throws TCException {
		// TODO Auto-generated method stub
		InterfaceAIFComponent ift = app.getTargetComponent();
		TCComponentBOMLine topbl = (TCComponentBOMLine) ift;
		TCComponentItemRevision boprev = null;
		try {
			boprev = topbl.getItemRevision();
		} catch (TCException e2) {
			// TODO Auto-generated catch block
			e2.printStackTrace();
		}
		// ��ȡ ��Ŀ-���� ��ѡ��
		projVehMap = ReportUtils.getDFL_Project_VehicleNo();
		String FamlilyCode = "";
		FamlilyCode = Util.getProperty(boprev, "project_ids");
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
		String procName = "00.����";


		// ������ʾ���Ȳ����ִ�в���
		ReportViwePanel viewPanel = new ReportViwePanel("���ɱ���");
		viewPanel.setVisible(true);

		viewPanel.addInfomation("��ʼ�������...\n", 10, 100);

		viewPanel.addInfomation("", 20, 100);
		// �жϱ���ģ���Ƿ�ά��
		// ��ѯ���浼��ģ��
		XSSFWorkbook book = null;

		viewPanel.addInfomation("", 40, 100);
		// �����ɷ��汨���ļ�
		String error = "";

		// ����BOP��������ȡ����������Ϣ
		String objectname = Util.getProperty(boprev, "object_name");

		String factoryline = ReportUtils.getFactoryLineByBOP(objectname);
		String factory = "";
		if(factoryline!=null && factoryline.length()>2) {
			factory = factoryline.substring(0, 3);
		}

		String[] cover = new String[5];
		cover[0] = "          ��    �ͣ�" + VehicleNo;
		cover[1] = "          ��    �Σ�" + Edition;
		cover[2] = "          �ļ���ţ�" + VehicleNo + "-" + factoryline + "-AB";
		cover[3] = "          �������ڣ�" + df.format(new Date());
		cover[4] = "          �������̣�" + factory + "������װ����";
		
		book = NewOutputDataToExcel.creatXSSFWorkbook(inputStream);

		String filename = procName;

		viewPanel.addInfomation("��ʼд���ݣ������ĵȴ�...\n", 60, 100);

		// ������·
		{
			Util.callByPass(session, true);
		}

		NewOutputDataToExcel.writeDataToSheet(book, cover);
		viewPanel.addInfomation("", 80, 100);
		NewOutputDataToExcel.exportFile(book, filename);

		String fullFileName = FileUtil.getReportFileName(filename);
		TCComponentDataset ds = Util.createDataset(session, filename, fullFileName, "MSExcelX", "excel");
		datasetList.add(ds);
		revlist.add(boprev);
		try {
			TCComponentItem document = ReportUtils.afterGenerateReportAction(datasetList, revlist, info, procName, "",
					session);
			// �����ļ���Ŵ��������ļ�����home��,������00.�����ļ��У��ѷ����ĵ��ŵ�00.�����ļ�����
			String topfoldername = VehicleNo + "-" + factoryline + "-AB";
			saveFileToFolder(document, topfoldername);

			// �ļ���ź��������
			TCProperty pdoc = document.getTCProperty("dfl9_vehiclePlant");
			if (pdoc != null) {
				if (pdoc != null) {
					pdoc.setStringValue(topfoldername);
					document.lock();
					document.save();
					document.unlock();
				}
			}
		} catch (TCException e) {
			e.printStackTrace();
			// EclipseUtils.info(Messages.FixtureOperation_15 + e.getMessage());
			return;
		}
		// �ر���·
		{
			Util.callByPass(session, false);
		}
		viewPanel.addInfomation("���������ɣ����ں�װ�������հ汾�����²鿴!", 100, 100);

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
			TCComponentFolderType foldertype = (TCComponentFolderType) session.getTypeComponent("Folder");
			if (folder == null) {
				folder = foldertype.create(topfoldername, "", "Folder");
				homefolder.add("contents", folder);
				childrenfolder = foldertype.create("00.����", "", "Folder");
				folder.add("contents", childrenfolder);
				childrenfolder.add("contents", document);
			} else {
				// ���ж��Ƿ��Ѿ�������00.�����ļ���
				AIFComponentContext[] icf1 = folder.getChildren();
				for (AIFComponentContext aif : icf1) {
					TCComponent tcc = (TCComponent) aif.getComponent();
					String obejctname = Util.getProperty(tcc, "object_name");
					if (tcc.getType().equals("Folder") && obejctname.equals("00.����")) {
						childrenfolder = (TCComponentFolder) tcc;
						break;
					}
				}
				if (childrenfolder == null) {
					childrenfolder = foldertype.create("00.����", "", "Folder");
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
			}

		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
