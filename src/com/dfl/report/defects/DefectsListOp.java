package com.dfl.report.defects;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

import javax.imageio.ImageIO;

import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dfl.report.util.FileUtil;
import com.dfl.report.util.OutputDataToExcel1;
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

public class DefectsListOp  {

	private AbstractAIFUIApplication app;
	private ArrayList numlist = new ArrayList();// ��ѡ���ֵ����
	private TCSession session;
	private TCComponentBOMLine topbomline;
	private ArrayList defectslist = new ArrayList();// ����δͨ���Ĳ��Խ������
	private ArrayList uplist = new ArrayList();
	private ArrayList downlist = new ArrayList();
	private ArrayList coverlist = new ArrayList();
	private ArrayList commonlist = new ArrayList();
	private HashMap<String, ArrayList> picmap = new HashMap<String, ArrayList>();
	private HashMap<String, ArrayList> picmap1 = new HashMap<String, ArrayList>();
	private HashMap<String, ArrayList> picmap2 = new HashMap<String, ArrayList>();
	private HashMap<String, ArrayList> picmap3 = new HashMap<String, ArrayList>();
	private HashMap<String, ArrayList> phomap = new HashMap<String, ArrayList>();
	private HashMap<String, ArrayList> phomap1 = new HashMap<String, ArrayList>();
	private HashMap<String, ArrayList> phomap2 = new HashMap<String, ArrayList>();
	private HashMap<String, ArrayList> phomap3 = new HashMap<String, ArrayList>();
	private String[] upvalue;// ������������
	private String[] downvalue;// ������������
	private String[] covervalue;// COVER��������
	private String[] commonvalue;// COMMON��������
	private int number = 0;// ������������
	private int number1 = 0;// ������������
	private int number2 = 0;// COVER��������
	private int number3 = 0;// COMMON��������
	private int pnumber;// ��״ͼƬ����
	private int phonumber;// Ҫ��ͼʾ����
	private String[][] data;// ��������
	private String[][] data1;// ��������
	private String[][] data2;// COVER����
	private String[][] data3;// COMMON����
	private String stage;// �׶�
	private TCComponent folder;
	private InterfaceAIFComponent[] aifComponents;
	
	SimpleDateFormat dateformat = new SimpleDateFormat("yyyyMMdd  HH");// �������ڸ�ʽ
	public DefectsListOp(AbstractAIFUIApplication app, String stage, ArrayList numlist,TCComponent folder,InterfaceAIFComponent[] aifComponents,TCSession session) {
		// TODO Auto-generated constructor stub
		this.app = app;
		this.numlist = numlist;
		this.stage=stage;
		this.folder = folder;
		this.aifComponents = aifComponents;
		this.session = session;
		initUI();
	}

	public void initUI() {
		// TODO Auto-generated method stub
		try {
			// ������ʾ���Ȳ����ִ�в���
			ReportViwePanel viewPanel = new ReportViwePanel("���ɱ���");
			viewPanel.setVisible(true);

			// ��ȡѡ��Ķ���
			//InterfaceAIFComponent[] aifComponents = app.getTargetComponents();
			TCComponentBOMLine bomline = (TCComponentBOMLine) aifComponents[0];
			topbomline = bomline.window().getTopBOMLine();

			viewPanel.addInfomation("���ڻ�ȡģ��...\n", 20, 100);
			// ��ѯ������ģ��
			InputStream inputStream = FileUtil.getTemplateFile("DFL_Template_DefectsCheckList");
			
			if (inputStream == null) {
				viewPanel.addInfomation("����û���ҵ�����Ҫ�����ߺ�һԪ���ģ�壬������TC�����ģ��(����Ϊ��DFL_Template_DefectsCheckList)\n", 100,100);
				return;
			}
			viewPanel.addInfomation("��ʼ�������...\n", 35, 100);
			// �Ȼ�ȡ���е�������
			List lineList = Util.getChildrenByParent(aifComponents);
			// ��ȡ���к�װ��λ���չ����Ĳ�������
			ArrayList list = new ArrayList();
			
			for(int i=0;i<lineList.size();i++) {
				TCComponentBOMLine pbl  = (TCComponentBOMLine) lineList.get(i);
				ArrayList childlist = Util.SearchTests(pbl);
				if(childlist!=null) {
					for(int j=0;j<childlist.size();j++) {
						if(!list.contains(childlist.get(j))) {
							list.add(childlist.get(j));
						}					
					}
				}
			}
			Comparator comparator = getComParator();

			if (list != null && list.size() > 0) {
				int index = 0;
				int pindex = 0;
				int index1 = 0;
				int pindex1 = 0;
				int index2 = 0;
				int pindex2 = 0;
				int index3 = 0;
				int pindex3 = 0;

				for (int i = 0; i < list.size(); i++) {
					TestManagerModelObject modelObject = (TestManagerModelObject) list.get(i);
					// ��ȡ��������ʵ��
					TCComponent testCaseInstance = modelObject.getTestComponent();
					// ��ȡ��������
					TCComponent testCase = modelObject.getTestCase();
					//��ȡҪ������
					String testcasetype=Util.getRelProperty(testCase, "b8_TestCaseType");
					//ȡ����Ҫ��������Ϊ0�Ĳ���������
					if(testcasetype.equals("0")) {
						if ((testCase != null && testCase instanceof TCComponentItem)
								&& (testCaseInstance != null && testCaseInstance instanceof TCComponentForm)) {

							// ��ȡ�����������°汾
							TCComponentItem testCaseItem = (TCComponentItem) testCase;
							TCComponentItemRevision testCaseRev = testCaseItem.getLatestItemRevision();

							// ��ȡ����ʵ�������������Ի��
							TCComponent[] activitys = testCaseInstance.getRelatedComponents("Tm0TestInstanceActivityRel");
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
									// ��Ҫ���򣨽�ȡ���µĲ��Ի��
									// //////////////////������ڽ�������
									Collections.sort(tempList2, comparator);

									// ȡtempList�е�һ��Ϊ�������ڵĲ��Ի
									TCComponentForm testactivity = (TCComponentForm) tempList2.get(0);

									// ��ȡ����δͨ���Ĳ��Խ��
									String defects = Util.getProperty(testactivity, "tm0ResultStatus");
									if (numlist.contains(defects)) {
										defectslist.add(testCaseRev);
										System.out.println("defectslist��" + defectslist);

										ArrayList piclist = new ArrayList();
										ArrayList pholist = new ArrayList();
										ArrayList piclist1 = new ArrayList();
										ArrayList pholist1 = new ArrayList();
										ArrayList piclist2 = new ArrayList();
										ArrayList pholist2 = new ArrayList();
										ArrayList piclist3 = new ArrayList();
										ArrayList pholist3 = new ArrayList();

//										// ����Ҫ�����У�����ʵ���ĸ������󣬻�ȡ������
//										TCComponent[] relcomps = testCaseItem.whereUsed(TCComponent.WHERE_USED_ALL);
//										if (relcomps != null && relcomps.length > 0) {
//											for (int j = 0; j < relcomps.length; j++) {
//												TCComponentItemRevision relcmp = (TCComponentItemRevision) relcomps[j];
//												// �жϣ�����/����/COVER/COMMON������
//												if (relcmp.isTypeOf("B8_RequirementRevision")) {
													String distinguish = Util.getProperty(testCaseRev, "b8_distinguish");// ����
													// ����
													if (distinguish.equals("UP")||distinguish.equals("����")) {

														// ��ȡ���Խ����ϵ�µ����ݼ�����
														TCComponent[] tdata = Util.getRelComponents(testactivity,"Tm0TestResultRel");
														//System.out.println("tdata  is :" + tdata.toString());
														// ��ͼƬ���͵����ݼ����ص�����
														for (int k = 0; k < tdata.length; k++) {
															File files = Util.downLoadPicture(tdata[k]);
															if (files != null) {
																piclist.add(files);
															}

														}
														pnumber = piclist.size();
														picmap.put(Integer.toString(index + 1), piclist);
														index++;

														// ��ȡ�ߺϹ�ϵ�µ����ݼ�����
														TCComponent[] pdata = Util.getRelComponents(testCaseRev,"B8_TestCase_Good");
														// ��ͼƬ���͵����ݼ����ص�����
														for (int m = 0; m < pdata.length; m++) {
															File pfiles = Util.downLoadPicture(pdata[m]);
															if (pfiles != null) {
																pholist.add(pfiles);
															}
														}
														phonumber = pholist.size();
														phomap.put(Integer.toString(pindex + 1), pholist);
														pindex++;

														upvalue = new String[11];
														upvalue[0] = "����";// ���̣��̶�ֵΪ�����塰
														upvalue[1] = "";// ���
														upvalue[2] = Util.getProperty(testactivity, "b8_TestStage");// �׶�
														upvalue[3] = "76000";// ��Ʒ����
														String bodyText = Util.getProperty(testCaseRev, "body_text");
														upvalue[4] = Util.getBodyText(bodyText);// Ҫ��
														upvalue[5] = Util.getProperty(testCaseRev, "b8_TestCaseID");// ����Ҫ�����
														//upvalue[6] = Util.getProperty(testCaseRev, "b8_Defects");// ���ߺ�����
														upvalue[6] = Util.getProperty(testactivity, "tm0Comment");// ���Խ������ ���ߺ�����
														upvalue[7] = Util.getProperty(testCaseRev, "object_name");// ����
														upvalue[8] = Util.getUserName(testactivity);// �����
														upvalue[9] = Util.getProperty(testactivity, "tm0ActivityDate");// ���ʱ��
														String level = Util.getProperty(testCaseRev, "b8_Level");// M/W
														if (level.equals("MUST")) {
															upvalue[10] = "M";
														}
														if (level.equals("WANT")) {
															upvalue[10] = "W";
														}
														uplist.add(upvalue);
														number++;
													}
													// ����
													if (distinguish.equals("DOWN")||distinguish.equals("����")) {

														// ��ȡ���Խ����ϵ�µ����ݼ�����
														TCComponent[] tdata1 = Util.getRelComponents(testactivity,
																"Tm0TestResultRel");
														// ��ͼƬ���͵����ݼ����ص�����
														for (int k = 0; k < tdata1.length; k++) {
															File files = Util.downLoadPicture(tdata1[k]);
															if (files != null) {
																piclist1.add(files);
															}
														}
														pnumber = piclist1.size();
														picmap1.put(Integer.toString(index1 + 1), piclist1);
														index1++;

														// ��ȡ�ߺϹ�ϵ�µ����ݼ�����
														TCComponent[] pdata1 = Util.getRelComponents(testCaseRev,
																"B8_TestCase_Good");
														// ��ͼƬ���͵����ݼ����ص�����
														for (int m = 0; m < pdata1.length; m++) {
															File pfiles1 = Util.downLoadPicture(pdata1[m]);
															if (pfiles1 != null) {
																pholist1.add(pfiles1);
															}
														}
														phonumber = pholist1.size();
														phomap1.put(Integer.toString(pindex1 + 1), pholist1);
														pindex1++;
														downvalue = new String[11];
														downvalue[0] = "����";// ���̣��̶�ֵΪ�����塰
														downvalue[1] = "";// ���
														downvalue[2] = Util.getProperty(testactivity, "b8_TestStage");// �׶�
														downvalue[3] = "";// ��Ʒ����
														String bodyText = Util.getProperty(testCaseRev, "body_text");
														downvalue[4] = Util.getBodyText(bodyText);// Ҫ��
														downvalue[5] = Util.getProperty(testCaseRev, "b8_TestCaseID");// ����Ҫ�����
														//downvalue[6] = Util.getProperty(testCaseRev, "b8_Defects");// ���ߺ�����
														downvalue[6] = Util.getProperty(testactivity, "tm0Comment");// ���Խ������ ���ߺ�����
														downvalue[7] = Util.getProperty(testCaseRev, "object_name");// ����
														downvalue[8] = Util.getUserName(testactivity);// �����
														downvalue[9] = Util.getProperty(testactivity, "tm0ActivityDate");// ���ʱ��
														String level = Util.getProperty(testCaseRev, "b8_Level");// M/W
														if (level.equals("MUST")) {
															downvalue[10] = "M";
														}
														if (level.equals("WANT")) {
															downvalue[10] = "W";
														}
														downlist.add(downvalue);
														number1++;
													}
													// COVER
													if (distinguish.equals("OVER")||distinguish.equals("COVER") ||distinguish.equals("Cover")) {

														// ��ȡ���Խ����ϵ�µ����ݼ�����
														TCComponent[] tdata2 = Util.getRelComponents(testactivity,
																"Tm0TestResultRel");
														// ��ͼƬ���͵����ݼ����ص�����
														for (int k = 0; k < tdata2.length; k++) {
															File files = Util.downLoadPicture(tdata2[k]);
															if (files != null) {
																piclist2.add(files);
															}
														}
														pnumber = piclist2.size();
														picmap2.put(Integer.toString(index2 + 1), piclist2);
														index2++;

														// ��ȡ�ߺϹ�ϵ�µ����ݼ�����
														TCComponent[] pdata2 = Util.getRelComponents(testCaseRev,
																"B8_TestCase_Good");
														// ��ͼƬ���͵����ݼ����ص�����
														for (int m = 0; m < pdata2.length; m++) {
															File pfiles = Util.downLoadPicture(pdata2[m]);
															if (pfiles != null) {
																pholist2.add(pfiles);
															}
														}
														phonumber = pholist2.size();
														phomap2.put(Integer.toString(pindex2 + 1), pholist2);
														pindex2++;

														covervalue = new String[11];
														covervalue[0] = "����";// ���̣��̶�ֵΪ�����塰
														covervalue[1] = "";// ���
														covervalue[2] = Util.getProperty(testactivity, "b8_TestStage");// �׶�
														covervalue[3] = "";// ��Ʒ����
														String bodyText = Util.getProperty(testCaseRev, "body_text");
														covervalue[4] = Util.getBodyText(bodyText);// Ҫ��
														covervalue[5] = Util.getProperty(testCaseRev, "b8_TestCaseID");// ����Ҫ�����
														//covervalue[6] = Util.getProperty(testCaseRev, "b8_Defects");// ���ߺ�����
														covervalue[6] = Util.getProperty(testactivity, "tm0Comment");// ���Խ������ ���ߺ�����
														covervalue[7] = Util.getProperty(testCaseRev, "object_name");// ����
														covervalue[8] = Util.getUserName(testactivity);// �����
														covervalue[9] = Util.getProperty(testactivity, "tm0ActivityDate");// ���ʱ��
														String level = Util.getProperty(testCaseRev, "b8_Level");// M/W
														if (level.equals("MUST")) {
															covervalue[10] = "M";
														}
														if (level.equals("WANT")) {
															covervalue[10] = "W";
														}
														coverlist.add(covervalue);
														number2++;
													}
													// COMMON
													if (distinguish.equals("COMMON")||distinguish.equals("Common")) {

														// ��ȡ���Խ����ϵ�µ����ݼ�����
														TCComponent[] tdata3 = Util.getRelComponents(testactivity,"Tm0TestResultRel");
														// ��ͼƬ���͵����ݼ����ص�����
														for (int k = 0; k < tdata3.length; k++) {
															File files = Util.downLoadPicture(tdata3[k]);
															if (files != null) {
																piclist3.add(files);
															}
														}
														pnumber = piclist3.size();
														picmap3.put(Integer.toString(index3 + 1), piclist3);
														index3++;

														// ��ȡ�ߺϹ�ϵ�µ����ݼ�����
														TCComponent[] pdata3 = Util.getRelComponents(testCaseRev,
																"B8_TestCase_Good");
														System.out.println("pdata  is :" + pdata3.toString());
														// ��ͼƬ���͵����ݼ����ص�����
														for (int m = 0; m < pdata3.length; m++) {
															File pfiles = Util.downLoadPicture(pdata3[m]);
															if (pfiles != null) {
																pholist3.add(pfiles);
															}

														}
														phonumber = pholist3.size();
														phomap3.put(Integer.toString(pindex3 + 1), pholist3);
														pindex3++;

														commonvalue = new String[11];
														commonvalue[0] = "����";// ���̣��̶�ֵΪ�����塰
														commonvalue[1] = "";// ���
														commonvalue[2] = Util.getProperty(testactivity, "b8_TestStage");// �׶�
														commonvalue[3] = "";// ��Ʒ����
														String bodyText = Util.getProperty(testCaseRev, "body_text");
														commonvalue[4] = Util.getBodyText(bodyText);// Ҫ��
														commonvalue[5] = Util.getProperty(testCaseRev, "b8_TestCaseID");// ����Ҫ�����
														//commonvalue[6] = Util.getProperty(testCaseRev, "b8_Defects");// ���ߺ�����
														commonvalue[6] = Util.getProperty(testactivity, "tm0Comment");// ���Խ������ ���ߺ�����
														commonvalue[7] = Util.getProperty(testCaseRev, "object_name");// ����
														commonvalue[8] = Util.getUserName(testactivity);// �����
														commonvalue[9] = Util.getProperty(testactivity, "tm0ActivityDate");// ���ʱ��
														String level = Util.getProperty(testCaseRev, "b8_Level");// M/W
														if (level.equals("MUST")) {
															commonvalue[10] = "M";
														}
														if (level.equals("WANT")) {
															commonvalue[10] = "W";
														}
														commonlist.add(commonvalue);
														number3++;
													}
												}
											}
										}
									}
								} else {
									System.out.println("error...");
								}
							}
						}
//					}
//				}
//			}

//			if(number==0&&number1==0&&number2==0&&number3==0) {
//				inputStream.close();
//				viewPanel.addInfomation("���棺��ȷ�Ϻ����빴ѡ���Խ���������Ӧ������Ҫ��\n", 100, 100);
//				return;
//			}else {
				// ���屨��Sheet1��������
				data = new String[(number * 10)][34];
				for (int i = 0; i < uplist.size(); i++) {
					for(int j=0;j<10;j++) {
						String[] str = (String[]) uplist.get(i);
						data[(i * 10)+j][0] = str[0];// ����
						data[(i * 10)+j][1] = Integer.toString(i + 1);// ���
						data[(i * 10)+j][2] = str[2];// �׶�
						data[(i * 10)+j][3] = str[7];// ����
						data[(i * 10)+j][4] = str[6];// ���ߺ�����
						data[(i * 10)+j][5] = str[3];// ��Ʒ����
						data[(i * 10)+j][6] = "";// ��״ͼʾ
						data[(i * 10)+j][7] = "";// ��״ͼʾ
						data[(i * 10)+j][8] = "";// ��״ͼʾ
						data[(i * 10)+j][9] = "";// ��״ͼʾ
						data[(i * 10)+j][10] = "";// ��״ͼʾ
						data[(i * 10)+j][11] = "";// ��״ͼʾ
						data[(i * 10)+j][12] = "";// ��״ͼʾ
						data[(i * 10)+j][13] = "";// ��״ͼʾ
						data[(i * 10)+j][14] = "";// ��״ͼʾ
						data[(i * 10)+j][15] = "";// ��״ͼʾ
						data[(i * 10)+j][16] = str[8];// �����
						data[(i * 10)+j][17] = str[9];// ���ʱ��
						data[(i * 10)+j][18] = str[4];// Ҫ��
						data[(i * 10)+j][19] = "";// Ҫ��ͼʾ
						data[(i * 10)+j][20] = str[10];// M/W
						data[(i * 10)+j][21] = "";// DNTC���ң�������
						data[(i * 10)+j][22] = "";// ��Ƶ���
						data[(i * 10)+j][23] = "";// �Բ�
						data[(i * 10)+j][24] = "";// �Բߵ���ʱ��
						data[(i * 10)+j][25] = "";// �Բ�ͼƬ
						data[(i * 10)+j][26] = "";// �᰸�ж�
						data[(i * 10)+j][27] = str[5];// ����Ҫ�����
						data[(i * 10)+j][28] = "";// ���ࣨ1��
						data[(i * 10)+j][29] = "";// ���ࣨ2��
						data[(i * 10)+j][30] = "";// ���
						data[(i * 10)+j][31] = "";// �ж�
						data[(i * 10)+j][32] = "";// �ж�Ҫ��
						data[(i * 10)+j][33] = "";// �ش����
					}


				}
				// ���屨����������
				data1 = new String[(number1 * 10)][34];
				for (int i = 0; i < downlist.size(); i++) {
					for(int j = 0; j < 10; j++) {
						String[] str = (String[]) downlist.get(i);
						data1[(i * 10)+j][0] = str[0];// ����
						data1[(i * 10)+j][1] = Integer.toString(i + 1);// ���
						data1[(i * 10)+j][2] = str[2];// �׶�
						data1[(i * 10)+j][3] = str[7];// ����
						data1[(i * 10)+j][4] = str[6];// ���ߺ�����
						data1[(i * 10)+j][5] = str[3];// ��Ʒ����
						data1[(i * 10)+j][6] = "";// ��״ͼʾ
						data1[(i * 10)+j][7] = "";// ��״ͼʾ
						data1[(i * 10)+j][8] = "";// ��״ͼʾ
						data1[(i * 10)+j][9] = "";// ��״ͼʾ
						data1[(i * 10)+j][10] = "";// ��״ͼʾ
						data1[(i * 10)+j][11] = "";// ��״ͼʾ
						data1[(i * 10)+j][12] = "";// ��״ͼʾ
						data1[(i * 10)+j][13] = "";// ��״ͼʾ
						data1[(i * 10)+j][14] = "";// ��״ͼʾ
						data1[(i * 10)+j][15] = "";// ��״ͼʾ
						data1[(i * 10)+j][16] = str[8];// �����
						data1[(i * 10)+j][17] = str[9];// ���ʱ��
						data1[(i * 10)+j][18] = str[4];// Ҫ��
						data1[(i * 10)+j][19] = "";// Ҫ��ͼʾ
						data1[(i * 10)+j][20] = str[10];// M/W
						data1[(i * 10)+j][21] = "";// DNTC���ң�������
						data1[(i * 10)+j][22] = "";// ��Ƶ���
						data1[(i * 10)+j][23] = "";// �Բ�
						data1[(i * 10)+j][24] = "";// �Բߵ���ʱ��
						data1[(i * 10)+j][25] = "";// �Բ�ͼƬ
						data1[(i * 10)+j][26] = "";// �᰸�ж�
						data1[(i * 10)+j][27] = str[5];// ����Ҫ�����
						data1[(i * 10)+j][28] = "";// ���ࣨ1��
						data1[(i * 10)+j][29] = "";// ���ࣨ2��
						data1[(i * 10)+j][30] = "";// ���
						data1[(i * 10)+j][31] = "";// �ж�
						data1[(i * 10)+j][32] = "";// �ж�Ҫ��
						data1[(i * 10)+j][33] = "";// �ش����
					}
				}
				// ���屨��Sheet��������
				data2 = new String[(number2 * 10)][34];
				for (int i = 0; i < coverlist.size(); i++) {
					for(int j = 0; j < 10; j++) {
						String[] str = (String[]) coverlist.get(i);
						data2[(i * 10)+j][0] = str[0];// ����
						data2[(i * 10)+j][1] = Integer.toString(i + 1);// ���
						data2[(i * 10)+j][2] = str[2];// �׶�
						data2[(i * 10)+j][3] = str[7];// ����
						data2[(i * 10)+j][4] = str[6];// ���ߺ�����
						data2[(i * 10)+j][5] = str[3];// ��Ʒ����
						data2[(i * 10)+j][6] = "";// ��״ͼʾ
						data2[(i * 10)+j][7] = "";// ��״ͼʾ
						data2[(i * 10)+j][8] = "";// ��״ͼʾ
						data2[(i * 10)+j][9] = "";// ��״ͼʾ
						data2[(i * 10)+j][10] = "";// ��״ͼʾ
						data2[(i * 10)+j][11] = "";// ��״ͼʾ
						data2[(i * 10)+j][12] = "";// ��״ͼʾ
						data2[(i * 10)+j][13] = "";// ��״ͼʾ
						data2[(i * 10)+j][14] = "";// ��״ͼʾ
						data2[(i * 10)+j][15] = "";// ��״ͼʾ
						data2[(i * 10)+j][16] = str[8];// �����
						data2[(i * 10)+j][17] = str[9];// ���ʱ��
						data2[(i * 10)+j][18] = str[4];// Ҫ��
						data2[(i * 10)+j][19] = "";// Ҫ��ͼʾ
						data2[(i * 10)+j][20] = str[10];// M/W
						data2[(i * 10)+j][21] = "";// DNTC���ң�������
						data2[(i * 10)+j][22] = "";// ��Ƶ���
						data2[(i * 10)+j][23] = "";// �Բ�
						data2[(i * 10)+j][24] = "";// �Բߵ���ʱ��
						data2[(i * 10)+j][25] = "";// �Բ�ͼƬ
						data2[(i * 10)+j][26] = "";// �᰸�ж�
						data2[(i * 10)+j][27] = str[5];// ����Ҫ�����
						data2[(i * 10)+j][28] = "";// ���ࣨ1��
						data2[(i * 10)+j][29] = "";// ���ࣨ2��
						data2[(i * 10)+j][30] = "";// ���
						data2[(i * 10)+j][31] = "";// �ж�
						data2[(i * 10)+j][32] = "";// �ж�Ҫ��
						data2[(i * 10)+j][33] = "";// �ش����
					}
				}
				// ���屨��Sheet��������
				data3 = new String[(number3 * 10)][34];
				for (int i = 0; i < commonlist.size(); i++) {
					for(int j = 0; j < 10; j++) {
						String[] str = (String[]) commonlist.get(i);
						data3[(i * 10)+j][0] = str[0];// ����
						data3[(i * 10)+j][1] = Integer.toString(i + 1);// ���
						data3[(i * 10)+j][2] = str[2];// �׶�
						data3[(i * 10)+j][3] = str[7];// ����
						data3[(i * 10)+j][4] = str[6];// ���ߺ�����
						data3[(i * 10)+j][5] = str[3];// ��Ʒ����
						data3[(i * 10)+j][6] = "";// ��״ͼʾ
						data3[(i * 10)+j][7] = "";// ��״ͼʾ
						data3[(i * 10)+j][8] = "";// ��״ͼʾ
						data3[(i * 10)+j][9] = "";// ��״ͼʾ
						data3[(i * 10)+j][10] = "";// ��״ͼʾ
						data3[(i * 10)+j][11] = "";// ��״ͼʾ
						data3[(i * 10)+j][12] = "";// ��״ͼʾ
						data3[(i * 10)+j][13] = "";// ��״ͼʾ
						data3[(i * 10)+j][14] = "";// ��״ͼʾ
						data3[(i * 10)+j][15] = "";// ��״ͼʾ
						data3[(i * 10)+j][16] = str[8];// �����
						data3[(i * 10)+j][17] = str[9];// ���ʱ��
						data3[(i * 10)+j][18] = str[4];// Ҫ��
						data3[(i * 10)+j][19] = "";// Ҫ��ͼʾ
						data3[(i * 10)+j][20] = str[10];// M/W
						data3[(i * 10)+j][21] = "";// DNTC���ң�������
						data3[(i * 10)+j][22] = "";// ��Ƶ���
						data3[(i * 10)+j][23] = "";// �Բ�
						data3[(i * 10)+j][24] = "";// �Բߵ���ʱ��
						data3[(i * 10)+j][25] = "";// �Բ�ͼƬ
						data3[(i * 10)+j][26] = "";// �᰸�ж�
						data3[(i * 10)+j][27] = str[5];// ����Ҫ�����
						data3[(i * 10)+j][28] = "";// ���ࣨ1��
						data3[(i * 10)+j][29] = "";// ���ࣨ2��
						data3[(i * 10)+j][30] = "";// ���
						data3[(i * 10)+j][31] = "";// �ж�
						data3[(i * 10)+j][32] = "";// �ж�Ҫ��
						data3[(i * 10)+j][33] = "";// �ش����
					}
				}
				// ����ģ�崴��Excelģ��
				XSSFWorkbook book = OutputDataToExcel3.creatXSSFWorkbook(inputStream, uplist, downlist, coverlist,commonlist);

				viewPanel.addInfomation("��ʼд���ݣ������ĵȴ�...\n", 70, 100);

				// д����
				if (uplist != null && uplist.size() > 0) {
					OutputDataToExcel3.writeDataToSheet1(book, data, picmap, phomap, true);
				}
				if (downlist != null && downlist.size() > 0) {
					OutputDataToExcel3.writeDataToSheet2(book, data1, picmap1, phomap1, true);
				}
				if (coverlist != null && coverlist.size() > 0) {
					OutputDataToExcel3.writeDataToSheet3(book, data2, picmap2, phomap2, true);
				}
				if (commonlist != null && commonlist.size() > 0) {
					OutputDataToExcel3.writeDataToSheet4(book, data3, picmap3, phomap3, true);
				}

				// ɾ����Sheetҳ
				if(!(uplist.isEmpty()&&downlist.isEmpty()&&coverlist.isEmpty()&&commonlist.isEmpty())) {
					OutputDataToExcel3.deleteXSSFWorkbook(book, uplist, downlist, coverlist, commonlist);
				}			
				// ������ļ�
				String familycode = topbomline.getItemRevision().getProperty("project_ids");// ����
				String vehicle = Util.getDFLProjectIdVehicle(familycode);
				if(vehicle==null || vehicle.isEmpty()) {
					vehicle = familycode;
				}
				String date = dateformat.format(new Date());
				// ������ļ�����
				String datasetname = vehicle+"����Ҫ�����ߺ�һԪ��("+stage+")"+"_"+date+"ʱ";
				String fileName = Util.formatString(datasetname);
				OutputDataToExcel3.exportFile(book, fileName.trim());
				//viewPanel.addInfomation("����������...\n", 100, 100);
				Util.saveFiles(fileName.trim(),datasetname, folder, session,"AI");

				viewPanel.addInfomation("���������ɣ�����ѡ�񱣴���ļ����²鿴��...\n", 100, 100);
//			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// ����
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
