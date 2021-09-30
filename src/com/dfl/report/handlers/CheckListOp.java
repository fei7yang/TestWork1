package com.dfl.report.handlers;

import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;

import com.dfl.report.util.FileUtil;
import com.dfl.report.util.OutputDataToExcel1;
import com.dfl.report.util.OutputDataToExcel3;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.SoaUtil;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentForm;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.testmanager.ui.model.TestManagerModelObject;
import com.teamcenter.rac.util.MessageBox;

public class CheckListOp {

	private AbstractAIFUIApplication app;
	private String stage;// �׶�
	private TCSession session;
	TCComponentBOMLine topbomline;
	private String[] value;// ��������
	private String[] value1;// Ҫ������������
	ArrayList valuelist = new ArrayList();// // ����ʵ�����Լ���
	ArrayList valuelist1 = new ArrayList();// // Ҫ�������Լ���
	int number = 0;// ��������
	String[][] data;
	private TCComponent folder;
	SimpleDateFormat dateformat = new SimpleDateFormat("yyyyMMdd  HH");// �������ڸ�ʽ
	SimpleDateFormat dateformat2 = new SimpleDateFormat("yyyy.MM.dd");// �������ڸ�ʽ
	Map<String, Integer[]> totalmap = new HashMap<String, Integer[]>();
	private InterfaceAIFComponent[] aifComponents;

	public CheckListOp(AbstractAIFUIApplication app, String stage, TCComponent folder, InterfaceAIFComponent[] aifComponents, TCSession session) {
		// TODO Auto-generated constructor stub
		this.app = app;
		this.stage = stage;
		this.folder = folder;
		this.session = session;
		this.aifComponents = aifComponents;
		initUI();
	}

	private void initUI() {
		// TODO Auto-generated method stub
		try {
			// ��ѯ������ģ��
			InputStream inputStream = FileUtil.getTemplateFile("DFL_Template_CheckList");
			System.out.println("inputStream=" + inputStream);
	        
			if (inputStream == null) {
				MessageBox.post("����û���ҵ�Ҫ������ģ�壬����ϵϵͳ����Ա��TC�����ģ��(����Ϊ��DFL_Template_CheckList)", "����",
						MessageBox.INFORMATION);
				return ;
			}
			
			// ������ʾ���Ȳ����ִ�в���
			ReportViwePanel viewPanel = new ReportViwePanel("���ɱ���");
			viewPanel.setVisible(true);

			// ��ȡѡ��Ķ���
			//InterfaceAIFComponent[] aifComponents = app.getTargetComponents();
			TCComponentBOMLine bomline = (TCComponentBOMLine) aifComponents[0];
			topbomline = bomline.window().getTopBOMLine();

			viewPanel.addInfomation("��ʼ�������...\n", 10, 100);
			

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

			if (list != null && list.size() > 0) {

				for (int i = 0; i < list.size(); i++) {
					TestManagerModelObject modelObject = (TestManagerModelObject) list.get(i);
					// ��ȡ��������ʵ��
					TCComponent testCaseInstance = modelObject.getTestComponent();
					// ��ȡ��������
					TCComponent testCase = modelObject.getTestCase();
					// ��ȡ��������������
					String testcasetype = Util.getRelProperty(testCase, "b8_TestCaseType");
					// ��ȡ����Ҫ��������Ϊ0��
					if (testcasetype.equals("0")) {
						if ((testCase != null && testCase instanceof TCComponentItem)
								&& (testCaseInstance != null && testCaseInstance instanceof TCComponentForm)) {

							// ��ȡ�����������°汾
							TCComponentItem testCaseItem = (TCComponentItem) testCase;
							TCComponentItemRevision testCaseRev = testCaseItem.getLatestItemRevision();
							String distinguish = Util.getProperty(testCaseRev, "b8_distinguish");// ����
							if (totalmap.containsKey(distinguish)) {
								Integer[] numbers = totalmap.get(distinguish);
								numbers[0] = numbers[0] + 1;
								totalmap.put(distinguish, numbers);
							} else {
								Integer[] numbers = new Integer[3];
								numbers[0] = 1;
								numbers[1] = 0;
								numbers[2] = 0;
								totalmap.put(distinguish, numbers);
							}
							// ��ȡ����ʵ�������������Ի��
							TCComponent[] activitys = testCaseInstance
									.getRelatedComponents("Tm0TestInstanceActivityRel");
							if (activitys != null && activitys.length > 0) {
								List tempList = new ArrayList();
								for (int j = 0; j < activitys.length; j++) {
									tempList.add(activitys[j]);
								}
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

									if (totalmap.containsKey(distinguish)) {
										Integer[] numbers = totalmap.get(distinguish);
										numbers[1] = numbers[1] + 1;
										totalmap.put(distinguish, numbers);
									}

									// ��Ҫ���򣨽�ȡ���µĲ��Ի��
									// //////////////////������ڽ�������
									Collections.sort(tempList2, comparator);

									// ȡtempList�е�һ��Ϊ�������ڵĲ��Ի
									TCComponentForm testactivity = (TCComponentForm) tempList2.get(0);
									// ��ʼ��ȡ��ͷ����
									value = new String[14];
									// ����Ҫ�����У�����ʵ���ĸ������󣬻�ȡ������
									TCComponent[] relcomps = testCaseItem.whereUsed(TCComponent.WHERE_USED_ALL);
									if (relcomps != null && relcomps.length > 0) {
										
										int relcompcount = 0;

										for (int j = 0; j < relcomps.length; j++) {

											TCComponentItemRevision relcmp = (TCComponentItemRevision) relcomps[j];

											if (relcmp.isTypeOf("B8_RequirementRevision")) {
												// ��ȡ��ͷ��2�к͵�8�е�����
												value1 = new String[2];
												value1[0] = Util.getProperty(relcmp, "object_name");// ��λ�����ƣ���������Ҫ������������
												value1[1] = distinguish;// ����
												valuelist1.add(value1);

												value[0] = "";// ģ���ͷǰ��һ��
												value[2] = Util.getProperty(testCaseRev, "b8_TestCaseID");// ����Ҫ�����
												// ��ȡ��λ���ţ�ȥ��
												String val = value[2].trim();
												String no = "";
//												if (val.contains("-")) {
//													no = val.substring(0, val.indexOf("-"));
//													if (no.startsWith("0") && no.length() > 1) {
//														no = no.substring(1, no.length());
//													}
//												}
												no = val.substring(0, 2);
												if (no.startsWith("0") && no.length() > 1) {
													no = no.substring(1, no.length());
												}

												value[1] = no;// ��λ���ţ�����Ҫ����ŵ�һ����-��ǰ���ֶΣ���λ0��ȥ��0

												value[3] = Util.getProperty(testCaseRev, "object_name");// ����Ҫ������
												value[4] = Util.getProperty(testCaseRev, "b8_Level");// MUST&WANT���֣�����
												value[5] = "";// ���ã�Ĭ��Ϊ�գ�
												value[6] = Util.getProperty(testCaseRev, "b8_PhaseIn");// ����ʱ��
												String status = Util.getProperty(testactivity, "tm0ResultStatus");// ȷ�Ͻ��:���������ԣ�����ϵͳ���еĲ��Խ��LOV,���ݱ�ͷ�����ɫ
												if (status != null && status.length() > 0) {
													value[7] = status.substring(0, 1);
												} else {
													value[7] = "";
												}
												if(relcompcount == 0) {
													if (value[7].trim().equals("1")) {
														if (totalmap.containsKey(distinguish)) {
															Integer[] numbers = totalmap.get(distinguish);
															numbers[2] = numbers[2] + 1;
															totalmap.put(distinguish, numbers);
														}
													}
												}												

												value[8] = Util.getProperty(testactivity, "tm0Comment");// �ж�����
												value[9] = Util.getUserName(testactivity);// ȷ����
												value[10] = Util.getProperty(testactivity, "tm0ActivityDate");// ȷ������(�����)
												value[11] = "";// ȷ�Ͻ��
												value[12] = "";// ȷ������
												value[13] = "";// ��ע
												valuelist.add(value);
												number++;
												
												relcompcount++ ;
											}
										}
									} else {
										value1 = new String[2];
										value1[0] = "";// ��λ�����ƣ���������Ҫ������������
										value1[1] = distinguish;// ����
										valuelist1.add(value1);
										value[0] = "";// ģ���ͷǰ��һ��
										value[2] = Util.getProperty(testCaseRev, "b8_TestCaseID");// ����Ҫ�����
										// ��ȡ��λ���ţ�ȥ��
										String val = value[2];
										String no = "";
										if (val.contains("-")) {
											no = val.substring(0, val.indexOf("-"));
											if (no.startsWith("0") && no.length() > 1) {
												no = no.substring(1, no.length());
											}
										}
										value[1] = no;// ��λ���ţ�����Ҫ����ŵ�һ����-��ǰ���ֶΣ���λ0��ȥ��0

										value[3] = Util.getProperty(testCaseRev, "object_name");// ����Ҫ������
										value[4] = Util.getProperty(testCaseRev, "b8_Level");// MUST&WANT���֣�����
										value[5] = "";// ���ã�Ĭ��Ϊ�գ�
										value[6] = Util.getProperty(testCaseRev, "b8_PhaseIn");// ����ʱ��
										String status = Util.getProperty(testactivity, "tm0ResultStatus");// ȷ�Ͻ��:���������ԣ�����ϵͳ���еĲ��Խ��LOV,���ݱ�ͷ�����ɫ
										if (status != null && status.length() > 0) {
											value[7] = status.substring(0, 1);
										} else {
											value[7] = "";
										}
										if (value[7].trim().equals("1")) {
											if (totalmap.containsKey(distinguish)) {
												Integer[] numbers = totalmap.get(distinguish);
												numbers[2] = numbers[2] + 1;
												totalmap.put(distinguish, numbers);
											}
										}
										value[8] = Util.getProperty(testactivity, "tm0Comment");// �ж�����
										value[9] = Util.getUserName(testactivity);// ȷ����
										value[10] = Util.getProperty(testactivity, "tm0ActivityDate");// ȷ������(�����)
										value[11] = "";// ȷ�Ͻ��
										value[12] = "";// ȷ������
										value[13] = "";// ��ע
										valuelist.add(value);
										number++;
									}
								}
							} else {
								System.out.println("error...");
							}
						}
					}
				}
			}

				viewPanel.addInfomation("�����������...\n", 35, 100);
				// ���屨��������
				data = new String[number][16];
				ArrayList datalist = new ArrayList();
				for (int i1 = 0; i1 < valuelist.size(); i1++) {
					String[] values = new String[17];
					String[] str1 = (String[]) valuelist1.get(i1);
					String[] str = (String[]) valuelist.get(i1);

					values[0] = str[0];// ģ���ͷǰ��һ��
					values[1] = str[1];// ��λ����
					values[2] = str1[0];// ��λ�����ƣ���������Ҫ������������
					values[3] = str[2];// ����Ҫ�����
					values[4] = str[3];// ����Ҫ������
					values[5] = str[4];// MUST&WANT���֣�����
					values[6] = str[5];// ����
					values[7] = str[6];//// ����ʱ��
					values[8] = str1[1];// ����
					values[9] = str[7];// ȷ�Ͻ��:���������ԣ����ݱ�ͷ�����ɫ
					values[10] = str[8];// �ж�����
					values[11] = str[9];// ȷ����
					values[12] = str[10];// ȷ������
					values[13] = "";
					values[14] = str[11];// ȷ�Ͻ��
					values[15] = str[12];// ȷ������
					values[16] = str[13];// ��ע
					datalist.add(values);
				}
				// ���ݲ�λ��������
				Comparator comparator2 = getComParatorBySerialID();
				Collections.sort(datalist, comparator2);
				String familycode = topbomline.getItemRevision().getProperty("project_ids");// ����
				String vehicle = Util.getDFLProjectIdVehicle(familycode);
				if(vehicle==null || vehicle.isEmpty()) {
					vehicle = familycode;
				}
				//ȥ��COMMON
				if(totalmap.containsKey("COMMON")) {
					totalmap.remove("COMMON");
				}
				ArrayList totallist = new ArrayList();
				totallist.add(totalmap);
				totallist.add(vehicle);
				totallist.add(dateformat2.format(new Date()));
				totallist.add(stage);
				
				viewPanel.addInfomation("��ʼд���ݣ������ĵȴ�...\n", 70, 100);
				// ������ļ�����

				String date = dateformat.format(new Date());
				String datasetname = vehicle + "����Ҫ������(" + stage + ")" + "_" + date + "ʱ";
				String fileName = Util.formatString(datasetname);
				OutputDataToExcel1 outdata = new OutputDataToExcel1(datalist, totallist, inputStream, fileName.trim());
				// ����ļ�
				// viewPanel.addInfomation("����������...\n", 100, 100);
				Util.saveFiles(fileName.trim(), datasetname, folder, session, "AP");

				viewPanel.addInfomation("���������ɣ�����ѡ�񱣴���ļ����²鿴��...\n", 100, 100);
			
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	private Comparator getComParatorBySerialID() {
		// TODO Auto-generated method stub
		Comparator comparator = new Comparator() {

			public int compare(Object obj, Object obj1) {
				Object[] comp1 = (Object[]) obj;
				Object[] comp2 = (Object[]) obj1;

				int d1 = 0;
				int d2 = 0;
				if (comp1[1] != null && !comp1[1].toString().isEmpty()) {
					d1 = Integer.parseInt(comp1[1].toString());
				}
				if (comp2[1] != null && !comp2[1].toString().isEmpty()) {
					d2 = Integer.parseInt(comp2[1].toString());
				}
				if (d2 > d1) {
					return -1;
				}

				return 1;
			}
		};

		return comparator;
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

}
