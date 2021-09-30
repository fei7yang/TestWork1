package com.dfl.report.handlers;

import java.awt.Container;
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
import com.dfl.report.util.NewOutputDataToExcel;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentForm;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.testmanager.ui.model.TestManagerModelObject;
import com.teamcenter.rac.util.MessageBox;

public class AntirustRequirementsCheckOp {

	private AbstractAIFUIApplication app;
	private Object[] value;
	private String[] value1;
	private String stage;
	SimpleDateFormat df = new SimpleDateFormat("yyyyMMdd  HH");// �������ڸ�ʽ
	private TCComponent folder;
	private InterfaceAIFComponent[] ifc;
	private TCSession session;

	public AntirustRequirementsCheckOp(AbstractAIFUIApplication app, String stage, TCComponent folder, InterfaceAIFComponent[] ifc, TCSession session) {
		// TODO Auto-generated constructor stub
		this.app = app;
		this.stage = stage;
		this.folder = folder;
		this.ifc = ifc;
		this.session= session;
		initUI();
	}

	private void initUI() {
		// TODO Auto-generated method stub
		try {
			ArrayList datalist = new ArrayList();// �������ݼ���

			//InterfaceAIFComponent[] target = app.getTargetComponents();
			TCComponentBOMLine firstbl = (TCComponentBOMLine) ifc[0];
			TCComponentBOMLine topbl = firstbl.window().getTopBOMLine();
			String familycode = Util.getProperty(topbl.getItemRevision(), "project_ids");
			String vecile = Util.getDFLProjectIdVehicle(familycode);
			if(vecile==null || vecile.isEmpty()) {
				vecile = familycode;
			}
			// ������ʾ���Ȳ����ִ�в���
			ReportViwePanel viewPanel = new ReportViwePanel("���ɱ���");
			viewPanel.setVisible(true);

			// viewPanel.addInfomation("��ʼ�������...\n", 5, 100);
			viewPanel.addInfomation("���ڻ�ȡģ��...\n", 10, 100);
			// ��ѯ����ģ��
			InputStream inputStream = FileUtil.getTemplateFile("DFL_Template_AntirustRequirementsCheck");

			if (inputStream == null) {
				viewPanel.addInfomation("����û���ҵ�����Ҫ������ģ�壬�������ģ��(����Ϊ��DFL_Template_AntirustRequirementsCheck)\n", 100,
						100);
				return;
			}

			// �Ȼ�ȡ���е�������
			List lineList = Util.getChildrenByParent(ifc);
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

			viewPanel.addInfomation("��ʼ�������...\n", 20, 100);

			if (list != null && list.size() > 0) {
				for (int i = 0; i < list.size(); i++) {
					TestManagerModelObject modelObject = (TestManagerModelObject) list.get(i);
					// ��ȡ��������ʵ��
					TCComponent testCaseInstance = modelObject.getTestComponent();

					// ��ȡ��������
					TCComponent testCase = modelObject.getTestCase();

					// ����Ҫ������ȡֵ
					String testCaseType = Util.getRelProperty(testCase, "b8_TestCaseType");

					// ȡҪ������Ϊ�����
					if (testCaseType.equals("2")) {
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
									value = new Object[15];
									// ����Ҫ�����У�����ʵ���ĸ������󣬻�ȡ������
									// TCComponent[] relcomps = testCaseItem.whereUsed(TCComponent.WHERE_USED_ALL);
									value[0] = Util.getProperty(testCaseRev, "b8_SerialID");// ���b8_SerialID
									value[1] = Util.getProperty(testCaseRev, "b8_distinguish");// ����
									value[2] = Util.getProperty(testCaseRev, "b8_TestCasePart");// ����b8_TestCasePart
									value[3] = Util.getProperty(testCaseRev, "b8_PastProblem");// ��ȥ�����
																								// b8_PastProblem
									value[4] = Util.getProperty(testCaseRev, "b8_Remarks");// ��עb8_Remarks

									File file1 = new File("Z:\\�����Eclipse����\\imges\\����ͼƬ1.png");
									TCComponent[] tccs = testCaseRev.getRelatedComponents("B8_TestCase_PartMap");
									ArrayList partmap = new ArrayList();
									for (TCComponent tcc : tccs) {
										File file = Util.downLoadPicture(tcc);
										if (file != null) {
											partmap.add(file);
										}
									}
									// partmap.add(file1);
									value[13] = Util.getProperty(testCaseRev, "b8_PointContent");// �㲿λͼ�ı�����
									value[5] = partmap;// �㲿λͼ ����ϵB8_TestCase_PartMap ��
									TCComponent[] tccs1 = testCaseRev.getRelatedComponents("B8_TestCase_Drawing");
									ArrayList partdrawing = new ArrayList();
									for (TCComponent tcc : tccs1) {
										File file = Util.downLoadPicture(tcc);
										if (file != null) {
											partdrawing.add(file);
										}
									}
									// partdrawing.add(file1);
									value[14] = "";// ͼָ��Ƿ����ı�����
									value[6] = partdrawing;// ͼָ��Ƿ�������ϵ B8_TestCase_Drawing��

									value[7] = Util.getProperty(testactivity, "b8_TestStage");// �׶�
									// value[7] = Util.getProperty(testCaseRev, "b8_PhaseIn");
									// �����������׶�
//									if(!stages.contains(value[7].toString())) {
//										stages.add(value[7].toString());
//									}
									value[8] = Util.getProperty(testactivity, "tm0ResultStatus");// �ж�
									value[9] = Util.getProperty(testactivity, "tm0Comment");// ����
									TCComponent[] tccs2 = testactivity.getRelatedComponents("Tm0TestResultRel");
									ArrayList comment = new ArrayList();
									for (TCComponent tcc : tccs2) {
										File file = Util.downLoadPicture(tcc);
										if (file != null) {
											comment.add(file);
										}
									}
									value[10] = comment;// ����ͼƬ
									value[11] = Util.getProperty(testCaseRev, "b8_distinguish");// ����
									value[12] = Util.getProperty(testactivity, "tm0ActivityDate");// �ճ�tm0ActivityDate
									datalist.add(value);
								}
							}
						}
					}

				}
			}

			viewPanel.addInfomation("��ʼд���ݣ������ĵȴ�...\n", 40, 100);

			Comparator comparator2 = getComParatorBySerialID();
			Collections.sort(datalist, comparator2);

			XSSFWorkbook book = NewOutputDataToExcel.creatXSSFWorkbook2(inputStream);

			// HashMap map = NewOutputDataToExcel.InitializeHeader(book,stages);

			viewPanel.addInfomation("", 60, 100);

			NewOutputDataToExcel.writeRequirementsDataToSheet(book, datalist);
			String date = df.format(new Date());
			String datasetname = vecile + "����Ҫ������" + "(" + stage + ")" + "_" + date + "ʱ";
			String filename = Util.formatString(datasetname);

			NewOutputDataToExcel.exportFile(book, filename);
			viewPanel.addInfomation("", 80, 100);

			// NewOutputDataToExcel.openFile(FileUtil.getReportFileName(filename.trim()));
			Util.saveFiles(filename, datasetname, folder, session, "AR");

			viewPanel.addInfomation("���������ɣ�����ѡ�񱣴���ļ����²鿴��\n", 100, 100);
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
