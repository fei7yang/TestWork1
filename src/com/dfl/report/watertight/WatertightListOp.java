package com.dfl.report.watertight;

import java.io.File;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dfl.report.util.FileUtil;
import com.dfl.report.util.OutputDataToExcel3;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentForm;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.testmanager.ui.model.TestManagerModelObject;

public class WatertightListOp {
	private AbstractAIFUIApplication app;
	private TCSession session;
	TCComponentBOMLine topbomline;
	private Object[] value;
	private String stage;// �׶�
	private TCComponent folder;
	private InterfaceAIFComponent[] aifComponents;

	SimpleDateFormat dateformat = new SimpleDateFormat("yyyyMMdd  HH");// �������ڸ�ʽ

	public WatertightListOp(AbstractAIFUIApplication app, String stage, TCComponent folder, InterfaceAIFComponent[] aifComponents, TCSession session) {
		// TODO Auto-generated constructor stub
		this.app = app;
		this.stage = stage;
		this.folder = folder;
		this.session = session;
		this.aifComponents = aifComponents;
		initUI();
	}

	public void initUI() {
		// TODO Auto-generated method stub
		try {
			ArrayList datalist = new ArrayList();// �������ݼ���
			// ������ʾ���Ȳ����ִ�в���
			ReportViwePanel viewPanel = new ReportViwePanel("���ɱ���");
			viewPanel.setVisible(true);

			// ��ȡѡ��Ķ���
			//InterfaceAIFComponent[] aifComponents = app.getTargetComponents();
			TCComponentBOMLine bomline = (TCComponentBOMLine) aifComponents[0];
			topbomline = bomline.window().getTopBOMLine();

			viewPanel.addInfomation("���ڻ�ȡģ��...\n", 30, 100);
			// ��ѯ������ģ��
			InputStream inputStream = FileUtil.getTemplateFile("DFL_Template_WatertightCheckList");

			if (inputStream == null) {

				/*
				 * viewPanel.addInfomation(
				 * "����û���ҵ�ˮ��Ҫ������ģ�壬������TC�����ģ��(����Ϊ��DFL_Template_WatertightCheckList)\n", 100,
				 * 100);
				 */
				System.out.println("����û���ҵ�ˮ��Ҫ������ģ�壬������TC�����ģ��(����Ϊ��DFL_Template_WatertightCheckList)");
				return;
			}

			// �Ȼ�ȡ���е�������
			List lineList = Util.getChildrenByParent(aifComponents);
			// ��ȡ���к�װ��λ���չ����Ĳ�������
			ArrayList list = new ArrayList();

			for (int i = 0; i < lineList.size(); i++) {
				TCComponentBOMLine pbl = (TCComponentBOMLine) lineList.get(i);
				ArrayList childlist = Util.SearchTests(pbl);
				if (childlist != null) {
					for (int j = 0; j < childlist.size(); j++) {
						if(!list.contains(childlist.get(j))) {
							list.add(childlist.get(j));
						}	
					}
				}
			}

			Comparator comparator = getComParator();

			// ˮ��Ҫ��λ��ͼƬ����ϵB8_TestCase_Watertight_Pos��
			ArrayList pngmap = new ArrayList();

			if (list != null && list.size() > 0) {
				for (int i = 0; i < list.size(); i++) {
					TestManagerModelObject modelObject = (TestManagerModelObject) list.get(i);
					// ��ȡ��������ʵ��
					TCComponent testCaseInstance = modelObject.getTestComponent();
					// ��ȡ��������
					TCComponent testCase = modelObject.getTestCase();
					// ��ȡ����Ϊ1�Ĳ���������ˮ��Ҫ����
					String testcasetype = Util.getRelProperty(testCase, "b8_TestCaseType");
					if (testcasetype.equals("1")) {
						if ((testCase != null && testCase instanceof TCComponentItem)
								&& (testCaseInstance != null && testCaseInstance instanceof TCComponentForm)) {
							// ��ȡ�����������°汾
							TCComponentItem testCaseItem = (TCComponentItem) testCase;
							TCComponentItemRevision testCaseRev = testCaseItem.getLatestItemRevision();

							// ��ȡ����ʵ�������������Ի��
							TCComponent[] activitys = testCaseInstance
									.getRelatedComponents("Tm0TestInstanceActivityRel");
							if (activitys != null && activitys.length > 0) {

								List tempList = new ArrayList();
								for (int j = 0; j < activitys.length; j++) {
									tempList.add(activitys[j]);
								}

								// Collections.sort(tempList, comparator);

								List tempList2 = new ArrayList();
								for (int k = 0; k < tempList.size(); k++) {
									// ȡtempList�в��Ի
									TCComponentForm testac = (TCComponentForm) tempList.get(k);

									// ���ѡ��Ľ׶�����
									String dqstage = Util.getProperty(testac, "b8_TestStage");

									if (dqstage.equals(stage)) {
										tempList2.add(testac);
									}

								}
								if (tempList2 != null && tempList2.size() > 0) {
									// ��Ҫ���򣨽�ȡ���µĲ��Ի��
									// //////////////////������ڽ�������
									Collections.sort(tempList2, comparator);

									// ȡtempList�в��Ի
									TCComponentForm testactivity = (TCComponentForm) tempList2.get(0);
									// ��ʼ��ȡ��ͷ����
									value = new Object[13];
									value[0] = Util.getProperty(testCaseRev, "b8_SerialID");// ���NO�����������
									value[1] = Util.getProperty(testCaseRev, "b8_distinguish");// ����
									value[2] = Util.getProperty(testCaseRev, "b8_TestCasePart");// ��λ
									value[3] = Util.getProperty(testCaseRev, "b8_ApplicableCar");// ���ù���
									value[4] = Util.getProperty(testCaseRev, "b8_DefectReason");// ���ߺ�Ҫ��
									// ˮ����Ҫע�ⲿλ��ͼƬ���ع�ϵB8_TestCase_Watertight
									ArrayList phomap = new ArrayList();
									// ��ȡB8_TestCase_Watertight��ϵ�µ����ݼ�����
									TCComponent[] tdata = testCaseRev.getRelatedComponents("B8_TestCase_Watertight");
									// ��ͼƬ���͵����ݼ����ص�����
									for (TCComponent tdt : tdata) {
										File file = Util.downLoadPicture(tdt);
										if (file != null) {
											phomap.add(file);
										}
									}

									value[5] = phomap;// ˮ����Ҫע�ⲿλͼƬ
									value[6] = Util.getProperty(testCaseRev, "b8_Check");// �����Ŀ
									value[7] = Util.getProperty(testCaseRev, "b8_Remarks");// ��ע
									value[8] = Util.getProperty(testactivity, "b8_TestStage");// �׶�
									String status = Util.getProperty(testactivity, "tm0ResultStatus");// ��񣨶�Ӧ���Խ����
									value[9] = status.substring(0, 1);// ��Ӧ���Խ��ǰһλ

									// ��ȡ���Խ����ϵ�µ����ݼ�����(��״ͼƬ)
									ArrayList picmap = new ArrayList();
									TCComponent[] tdata1 = testactivity.getRelatedComponents("Tm0TestResultRel");
									// ��ͼƬ���͵����ݼ����ص�����
									for (TCComponent tdt1 : tdata1) {
										File file = Util.downLoadPicture(tdt1);
										if (file != null) {
											picmap.add(file);
										}
									}
									value[10] = picmap;// ��״ͼʾ

									value[11] = Util.getProperty(testactivity, "tm0Comment");// ��������

									TCComponent[] tdata2 = testCaseRev
											.getRelatedComponents("B8_TestCase_Watertight_Pos");
									// ��ͼƬ���͵����ݼ����ص�����
									for (TCComponent tdt2 : tdata2) {
										File file = Util.downLoadPicture(tdt2);
										if (file != null) {
											if (pngmap.size() < 1) {
												pngmap.add(file);
											}
										}
									}
									value[12] = pngmap;// ˮ��Ҫ��λ��ͼƬ

									datalist.add(value);
								}
							}
						}
					}
				}
			}
			viewPanel.addInfomation("��ʼд���ݣ������ĵȴ�...\n", 50, 100);

			Comparator comparator2 = getComParatorBySerialID();
			Collections.sort(datalist, comparator2);
			// ����ģ�崴��Excelģ��
			XSSFWorkbook book = OutputDataToExcel3.creatXSSFWorkbook(inputStream);

			viewPanel.addInfomation("", 80, 100);
			// д����
			OutputDataToExcel3.writeDataToSheet(book, datalist);

			// ����ļ�
			String familycode = topbomline.getItemRevision().getProperty("project_ids");// ����
			String vehicle = Util.getDFLProjectIdVehicle(familycode);
			if(vehicle==null || vehicle.isEmpty()) {
				vehicle = familycode;
			}
			String date = dateformat.format(new Date());
			// ������ļ�����
			String datasetname = vehicle + "ˮ��Ҫ�������һԪ��(" + stage + ")" + "_" + date + "ʱ";
			String fileName = Util.formatString(datasetname);
			OutputDataToExcel3.exportFile(book, fileName.trim());
			// viewPanel.addInfomation("����������...\n", 100, 100);
			Util.saveFiles(fileName.trim(), datasetname, folder, session, "AQ");

			viewPanel.addInfomation("���������ɣ�����ѡ�񱣴���ļ����²鿴��...\n", 100, 100);
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	private Comparator getComParator() {
		// TODO Auto-generated method stub
		Comparator comparator = new Comparator() {

			public int compare(Object obj, Object obj1) {
				TCComponent comp1 = (TCComponent) obj;
				TCComponent comp2 = (TCComponent) obj1;

				try {
					Date d1 = comp1.getDateProperty("tm0ActivityDate");
					Date d2 = comp2.getDateProperty("tm0ActivityDate");
					return d2.compareTo(d1);
				} catch (TCException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

				return -1;
			}
		};

		return comparator;
	}

	private Comparator getComParatorBySerialID() {
		// TODO Auto-generated method stub
		Comparator comparator = new Comparator() {

			public int compare(Object obj, Object obj1) {
				Object[] comp1 = (Object[]) obj;
				Object[] comp2 = (Object[]) obj1;

				int d1 = 0;
				int d2 = 0;
				if (comp1[0] != null && !comp1[0].toString().isEmpty()) {
					d1 = Integer.parseInt(comp1[0].toString());
				}
				if (comp2[0] != null && !comp2[0].toString().isEmpty()) {
					d2 = Integer.parseInt(comp2[0].toString());
				}
				if (d2 > d1) {
					return -1;
				}

				return 1;
			}
		};

		return comparator;
	}

}
