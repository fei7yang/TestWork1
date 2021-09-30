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
	private ArrayList numlist = new ArrayList();// 复选框的值集合
	private TCSession session;
	private TCComponentBOMLine topbomline;
	private ArrayList defectslist = new ArrayList();// 所有未通过的测试结果集合
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
	private String[] upvalue;// 上屋属性数组
	private String[] downvalue;// 下屋属性数组
	private String[] covervalue;// COVER属性数组
	private String[] commonvalue;// COMMON属性数组
	private int number = 0;// 上屋数据行数
	private int number1 = 0;// 下屋数据行数
	private int number2 = 0;// COVER数据行数
	private int number3 = 0;// COMMON数据行数
	private int pnumber;// 现状图片数量
	private int phonumber;// 要望图示数量
	private String[][] data;// 上屋数据
	private String[][] data1;// 下屋数据
	private String[][] data2;// COVER数据
	private String[][] data3;// COMMON数据
	private String stage;// 阶段
	private TCComponent folder;
	private InterfaceAIFComponent[] aifComponents;
	
	SimpleDateFormat dateformat = new SimpleDateFormat("yyyyMMdd  HH");// 设置日期格式
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
			// 界面显示进度并输出执行步骤
			ReportViwePanel viewPanel = new ReportViwePanel("生成报表");
			viewPanel.setVisible(true);

			// 获取选择的对象
			//InterfaceAIFComponent[] aifComponents = app.getTargetComponents();
			TCComponentBOMLine bomline = (TCComponentBOMLine) aifComponents[0];
			topbomline = bomline.window().getTopBOMLine();

			viewPanel.addInfomation("正在获取模板...\n", 20, 100);
			// 查询并导出模板
			InputStream inputStream = FileUtil.getTemplateFile("DFL_Template_DefectsCheckList");
			
			if (inputStream == null) {
				viewPanel.addInfomation("错误：没有找到生产要件不具合一元表的模板，请先在TC中添加模板(名称为：DFL_Template_DefectsCheckList)\n", 100,100);
				return;
			}
			viewPanel.addInfomation("开始输出报表...\n", 35, 100);
			// 先获取所有的虚层产线
			List lineList = Util.getChildrenByParent(aifComponents);
			// 获取所有焊装工位工艺关联的测试用例
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
					// 获取测试用例实例
					TCComponent testCaseInstance = modelObject.getTestComponent();
					// 获取测试用例
					TCComponent testCase = modelObject.getTestCase();
					//获取要件类型
					String testcasetype=Util.getRelProperty(testCase, "b8_TestCaseType");
					//取生产要件（类型为0的测试用例）
					if(testcasetype.equals("0")) {
						if ((testCase != null && testCase instanceof TCComponentItem)
								&& (testCaseInstance != null && testCaseInstance instanceof TCComponentForm)) {

							// 获取测试用例最新版本
							TCComponentItem testCaseItem = (TCComponentItem) testCase;
							TCComponentItemRevision testCaseRev = testCaseItem.getLatestItemRevision();

							// 获取测试实例活动对象表单（测试活动）
							TCComponent[] activitys = testCaseInstance.getRelatedComponents("Tm0TestInstanceActivityRel");
							if (activitys != null && activitys.length > 0) {

								List tempList = new ArrayList();
								for (int j = 0; j < activitys.length; j++) {
									tempList.add(activitys[j]);
								}

								List tempList2 = new ArrayList();
								for (int k = 0; k < tempList.size(); k++) {
									// 取tempList中测试活动
									TCComponentForm testac = (TCComponentForm) tempList.get(k);

									// 输出选择的阶段数据
									String dqstage = Util.getProperty(testac, "b8_TestStage");

									if (dqstage.equals(stage)) {
										tempList2.add(testac);
									}

								}
								if (tempList2 != null && tempList2.size() > 0) {
									// 需要排序（仅取最新的测试活动）
									// //////////////////按活动日期进行排序
									Collections.sort(tempList2, comparator);

									// 取tempList中第一个为最新日期的测试活动
									TCComponentForm testactivity = (TCComponentForm) tempList2.get(0);

									// 获取所有未通过的测试结果
									String defects = Util.getProperty(testactivity, "tm0ResultStatus");
									if (numlist.contains(defects)) {
										defectslist.add(testCaseRev);
										System.out.println("defectslist：" + defectslist);

										ArrayList piclist = new ArrayList();
										ArrayList pholist = new ArrayList();
										ArrayList piclist1 = new ArrayList();
										ArrayList pholist1 = new ArrayList();
										ArrayList piclist2 = new ArrayList();
										ArrayList pholist2 = new ArrayList();
										ArrayList piclist3 = new ArrayList();
										ArrayList pholist3 = new ArrayList();

//										// 查找要件库中，测试实例的父级对象，获取其属性
//										TCComponent[] relcomps = testCaseItem.whereUsed(TCComponent.WHERE_USED_ALL);
//										if (relcomps != null && relcomps.length > 0) {
//											for (int j = 0; j < relcomps.length; j++) {
//												TCComponentItemRevision relcmp = (TCComponentItemRevision) relcomps[j];
//												// 判断（上屋/下屋/COVER/COMMON）类型
//												if (relcmp.isTypeOf("B8_RequirementRevision")) {
													String distinguish = Util.getProperty(testCaseRev, "b8_distinguish");// 区分
													// 上屋
													if (distinguish.equals("UP")||distinguish.equals("上屋")) {

														// 获取测试结果关系下的数据集对象
														TCComponent[] tdata = Util.getRelComponents(testactivity,"Tm0TestResultRel");
														//System.out.println("tdata  is :" + tdata.toString());
														// 将图片类型的数据集下载到本地
														for (int k = 0; k < tdata.length; k++) {
															File files = Util.downLoadPicture(tdata[k]);
															if (files != null) {
																piclist.add(files);
															}

														}
														pnumber = piclist.size();
														picmap.put(Integer.toString(index + 1), piclist);
														index++;

														// 获取具合关系下的数据集对象
														TCComponent[] pdata = Util.getRelComponents(testCaseRev,"B8_TestCase_Good");
														// 将图片类型的数据集下载到本地
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
														upvalue[0] = "车体";// 工程，固定值为”车体“
														upvalue[1] = "";// 序号
														upvalue[2] = Util.getProperty(testactivity, "b8_TestStage");// 阶段
														upvalue[3] = "76000";// 部品番号
														String bodyText = Util.getProperty(testCaseRev, "body_text");
														upvalue[4] = Util.getBodyText(bodyText);// 要望
														upvalue[5] = Util.getProperty(testCaseRev, "b8_TestCaseID");// 生产要件编号
														//upvalue[6] = Util.getProperty(testCaseRev, "b8_Defects");// 不具合内容
														upvalue[6] = Util.getProperty(testactivity, "tm0Comment");// 测试结果描述 不具合内容
														upvalue[7] = Util.getProperty(testCaseRev, "object_name");// 件名
														upvalue[8] = Util.getUserName(testactivity);// 提出人
														upvalue[9] = Util.getProperty(testactivity, "tm0ActivityDate");// 提出时间
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
													// 下屋
													if (distinguish.equals("DOWN")||distinguish.equals("下屋")) {

														// 获取测试结果关系下的数据集对象
														TCComponent[] tdata1 = Util.getRelComponents(testactivity,
																"Tm0TestResultRel");
														// 将图片类型的数据集下载到本地
														for (int k = 0; k < tdata1.length; k++) {
															File files = Util.downLoadPicture(tdata1[k]);
															if (files != null) {
																piclist1.add(files);
															}
														}
														pnumber = piclist1.size();
														picmap1.put(Integer.toString(index1 + 1), piclist1);
														index1++;

														// 获取具合关系下的数据集对象
														TCComponent[] pdata1 = Util.getRelComponents(testCaseRev,
																"B8_TestCase_Good");
														// 将图片类型的数据集下载到本地
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
														downvalue[0] = "车体";// 工程，固定值为”车体“
														downvalue[1] = "";// 序号
														downvalue[2] = Util.getProperty(testactivity, "b8_TestStage");// 阶段
														downvalue[3] = "";// 部品番号
														String bodyText = Util.getProperty(testCaseRev, "body_text");
														downvalue[4] = Util.getBodyText(bodyText);// 要望
														downvalue[5] = Util.getProperty(testCaseRev, "b8_TestCaseID");// 生产要件编号
														//downvalue[6] = Util.getProperty(testCaseRev, "b8_Defects");// 不具合内容
														downvalue[6] = Util.getProperty(testactivity, "tm0Comment");// 测试结果描述 不具合内容
														downvalue[7] = Util.getProperty(testCaseRev, "object_name");// 件名
														downvalue[8] = Util.getUserName(testactivity);// 提出人
														downvalue[9] = Util.getProperty(testactivity, "tm0ActivityDate");// 提出时间
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

														// 获取测试结果关系下的数据集对象
														TCComponent[] tdata2 = Util.getRelComponents(testactivity,
																"Tm0TestResultRel");
														// 将图片类型的数据集下载到本地
														for (int k = 0; k < tdata2.length; k++) {
															File files = Util.downLoadPicture(tdata2[k]);
															if (files != null) {
																piclist2.add(files);
															}
														}
														pnumber = piclist2.size();
														picmap2.put(Integer.toString(index2 + 1), piclist2);
														index2++;

														// 获取具合关系下的数据集对象
														TCComponent[] pdata2 = Util.getRelComponents(testCaseRev,
																"B8_TestCase_Good");
														// 将图片类型的数据集下载到本地
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
														covervalue[0] = "车体";// 工程，固定值为”车体“
														covervalue[1] = "";// 序号
														covervalue[2] = Util.getProperty(testactivity, "b8_TestStage");// 阶段
														covervalue[3] = "";// 部品番号
														String bodyText = Util.getProperty(testCaseRev, "body_text");
														covervalue[4] = Util.getBodyText(bodyText);// 要望
														covervalue[5] = Util.getProperty(testCaseRev, "b8_TestCaseID");// 生产要件编号
														//covervalue[6] = Util.getProperty(testCaseRev, "b8_Defects");// 不具合内容
														covervalue[6] = Util.getProperty(testactivity, "tm0Comment");// 测试结果描述 不具合内容
														covervalue[7] = Util.getProperty(testCaseRev, "object_name");// 件名
														covervalue[8] = Util.getUserName(testactivity);// 提出人
														covervalue[9] = Util.getProperty(testactivity, "tm0ActivityDate");// 提出时间
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

														// 获取测试结果关系下的数据集对象
														TCComponent[] tdata3 = Util.getRelComponents(testactivity,"Tm0TestResultRel");
														// 将图片类型的数据集下载到本地
														for (int k = 0; k < tdata3.length; k++) {
															File files = Util.downLoadPicture(tdata3[k]);
															if (files != null) {
																piclist3.add(files);
															}
														}
														pnumber = piclist3.size();
														picmap3.put(Integer.toString(index3 + 1), piclist3);
														index3++;

														// 获取具合关系下的数据集对象
														TCComponent[] pdata3 = Util.getRelComponents(testCaseRev,
																"B8_TestCase_Good");
														System.out.println("pdata  is :" + pdata3.toString());
														// 将图片类型的数据集下载到本地
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
														commonvalue[0] = "车体";// 工程，固定值为”车体“
														commonvalue[1] = "";// 序号
														commonvalue[2] = Util.getProperty(testactivity, "b8_TestStage");// 阶段
														commonvalue[3] = "";// 部品番号
														String bodyText = Util.getProperty(testCaseRev, "body_text");
														commonvalue[4] = Util.getBodyText(bodyText);// 要望
														commonvalue[5] = Util.getProperty(testCaseRev, "b8_TestCaseID");// 生产要件编号
														//commonvalue[6] = Util.getProperty(testCaseRev, "b8_Defects");// 不具合内容
														commonvalue[6] = Util.getProperty(testactivity, "tm0Comment");// 测试结果描述 不具合内容
														commonvalue[7] = Util.getProperty(testCaseRev, "object_name");// 件名
														commonvalue[8] = Util.getUserName(testactivity);// 提出人
														commonvalue[9] = Util.getProperty(testactivity, "tm0ActivityDate");// 提出时间
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
//				viewPanel.addInfomation("警告：请确认含有与勾选测试结果类型相对应的生产要件\n", 100, 100);
//				return;
//			}else {
				// 定义报表Sheet1中行数据
				data = new String[(number * 10)][34];
				for (int i = 0; i < uplist.size(); i++) {
					for(int j=0;j<10;j++) {
						String[] str = (String[]) uplist.get(i);
						data[(i * 10)+j][0] = str[0];// 工程
						data[(i * 10)+j][1] = Integer.toString(i + 1);// 序号
						data[(i * 10)+j][2] = str[2];// 阶段
						data[(i * 10)+j][3] = str[7];// 件名
						data[(i * 10)+j][4] = str[6];// 不具合内容
						data[(i * 10)+j][5] = str[3];// 部品番号
						data[(i * 10)+j][6] = "";// 现状图示
						data[(i * 10)+j][7] = "";// 现状图示
						data[(i * 10)+j][8] = "";// 现状图示
						data[(i * 10)+j][9] = "";// 现状图示
						data[(i * 10)+j][10] = "";// 现状图示
						data[(i * 10)+j][11] = "";// 现状图示
						data[(i * 10)+j][12] = "";// 现状图示
						data[(i * 10)+j][13] = "";// 现状图示
						data[(i * 10)+j][14] = "";// 现状图示
						data[(i * 10)+j][15] = "";// 现状图示
						data[(i * 10)+j][16] = str[8];// 提出人
						data[(i * 10)+j][17] = str[9];// 提出时间
						data[(i * 10)+j][18] = str[4];// 要望
						data[(i * 10)+j][19] = "";// 要望图示
						data[(i * 10)+j][20] = str[10];// M/W
						data[(i * 10)+j][21] = "";// DNTC科室（主担）
						data[(i * 10)+j][22] = "";// 设计担当
						data[(i * 10)+j][23] = "";// 对策
						data[(i * 10)+j][24] = "";// 对策导入时间
						data[(i * 10)+j][25] = "";// 对策图片
						data[(i * 10)+j][26] = "";// 结案判断
						data[(i * 10)+j][27] = str[5];// 生产要件编号
						data[(i * 10)+j][28] = "";// 分类（1）
						data[(i * 10)+j][29] = "";// 分类（2）
						data[(i * 10)+j][30] = "";// 版次
						data[(i * 10)+j][31] = "";// 判断
						data[(i * 10)+j][32] = "";// 判断要望
						data[(i * 10)+j][33] = "";// 重大课题
					}


				}
				// 定义报表中行数据
				data1 = new String[(number1 * 10)][34];
				for (int i = 0; i < downlist.size(); i++) {
					for(int j = 0; j < 10; j++) {
						String[] str = (String[]) downlist.get(i);
						data1[(i * 10)+j][0] = str[0];// 工程
						data1[(i * 10)+j][1] = Integer.toString(i + 1);// 序号
						data1[(i * 10)+j][2] = str[2];// 阶段
						data1[(i * 10)+j][3] = str[7];// 件名
						data1[(i * 10)+j][4] = str[6];// 不具合内容
						data1[(i * 10)+j][5] = str[3];// 部品番号
						data1[(i * 10)+j][6] = "";// 现状图示
						data1[(i * 10)+j][7] = "";// 现状图示
						data1[(i * 10)+j][8] = "";// 现状图示
						data1[(i * 10)+j][9] = "";// 现状图示
						data1[(i * 10)+j][10] = "";// 现状图示
						data1[(i * 10)+j][11] = "";// 现状图示
						data1[(i * 10)+j][12] = "";// 现状图示
						data1[(i * 10)+j][13] = "";// 现状图示
						data1[(i * 10)+j][14] = "";// 现状图示
						data1[(i * 10)+j][15] = "";// 现状图示
						data1[(i * 10)+j][16] = str[8];// 提出人
						data1[(i * 10)+j][17] = str[9];// 提出时间
						data1[(i * 10)+j][18] = str[4];// 要望
						data1[(i * 10)+j][19] = "";// 要望图示
						data1[(i * 10)+j][20] = str[10];// M/W
						data1[(i * 10)+j][21] = "";// DNTC科室（主担）
						data1[(i * 10)+j][22] = "";// 设计担当
						data1[(i * 10)+j][23] = "";// 对策
						data1[(i * 10)+j][24] = "";// 对策导入时间
						data1[(i * 10)+j][25] = "";// 对策图片
						data1[(i * 10)+j][26] = "";// 结案判断
						data1[(i * 10)+j][27] = str[5];// 生产要件编号
						data1[(i * 10)+j][28] = "";// 分类（1）
						data1[(i * 10)+j][29] = "";// 分类（2）
						data1[(i * 10)+j][30] = "";// 版次
						data1[(i * 10)+j][31] = "";// 判断
						data1[(i * 10)+j][32] = "";// 判断要望
						data1[(i * 10)+j][33] = "";// 重大课题
					}
				}
				// 定义报表Sheet中行数据
				data2 = new String[(number2 * 10)][34];
				for (int i = 0; i < coverlist.size(); i++) {
					for(int j = 0; j < 10; j++) {
						String[] str = (String[]) coverlist.get(i);
						data2[(i * 10)+j][0] = str[0];// 工程
						data2[(i * 10)+j][1] = Integer.toString(i + 1);// 序号
						data2[(i * 10)+j][2] = str[2];// 阶段
						data2[(i * 10)+j][3] = str[7];// 件名
						data2[(i * 10)+j][4] = str[6];// 不具合内容
						data2[(i * 10)+j][5] = str[3];// 部品番号
						data2[(i * 10)+j][6] = "";// 现状图示
						data2[(i * 10)+j][7] = "";// 现状图示
						data2[(i * 10)+j][8] = "";// 现状图示
						data2[(i * 10)+j][9] = "";// 现状图示
						data2[(i * 10)+j][10] = "";// 现状图示
						data2[(i * 10)+j][11] = "";// 现状图示
						data2[(i * 10)+j][12] = "";// 现状图示
						data2[(i * 10)+j][13] = "";// 现状图示
						data2[(i * 10)+j][14] = "";// 现状图示
						data2[(i * 10)+j][15] = "";// 现状图示
						data2[(i * 10)+j][16] = str[8];// 提出人
						data2[(i * 10)+j][17] = str[9];// 提出时间
						data2[(i * 10)+j][18] = str[4];// 要望
						data2[(i * 10)+j][19] = "";// 要望图示
						data2[(i * 10)+j][20] = str[10];// M/W
						data2[(i * 10)+j][21] = "";// DNTC科室（主担）
						data2[(i * 10)+j][22] = "";// 设计担当
						data2[(i * 10)+j][23] = "";// 对策
						data2[(i * 10)+j][24] = "";// 对策导入时间
						data2[(i * 10)+j][25] = "";// 对策图片
						data2[(i * 10)+j][26] = "";// 结案判断
						data2[(i * 10)+j][27] = str[5];// 生产要件编号
						data2[(i * 10)+j][28] = "";// 分类（1）
						data2[(i * 10)+j][29] = "";// 分类（2）
						data2[(i * 10)+j][30] = "";// 版次
						data2[(i * 10)+j][31] = "";// 判断
						data2[(i * 10)+j][32] = "";// 判断要望
						data2[(i * 10)+j][33] = "";// 重大课题
					}
				}
				// 定义报表Sheet中行数据
				data3 = new String[(number3 * 10)][34];
				for (int i = 0; i < commonlist.size(); i++) {
					for(int j = 0; j < 10; j++) {
						String[] str = (String[]) commonlist.get(i);
						data3[(i * 10)+j][0] = str[0];// 工程
						data3[(i * 10)+j][1] = Integer.toString(i + 1);// 序号
						data3[(i * 10)+j][2] = str[2];// 阶段
						data3[(i * 10)+j][3] = str[7];// 件名
						data3[(i * 10)+j][4] = str[6];// 不具合内容
						data3[(i * 10)+j][5] = str[3];// 部品番号
						data3[(i * 10)+j][6] = "";// 现状图示
						data3[(i * 10)+j][7] = "";// 现状图示
						data3[(i * 10)+j][8] = "";// 现状图示
						data3[(i * 10)+j][9] = "";// 现状图示
						data3[(i * 10)+j][10] = "";// 现状图示
						data3[(i * 10)+j][11] = "";// 现状图示
						data3[(i * 10)+j][12] = "";// 现状图示
						data3[(i * 10)+j][13] = "";// 现状图示
						data3[(i * 10)+j][14] = "";// 现状图示
						data3[(i * 10)+j][15] = "";// 现状图示
						data3[(i * 10)+j][16] = str[8];// 提出人
						data3[(i * 10)+j][17] = str[9];// 提出时间
						data3[(i * 10)+j][18] = str[4];// 要望
						data3[(i * 10)+j][19] = "";// 要望图示
						data3[(i * 10)+j][20] = str[10];// M/W
						data3[(i * 10)+j][21] = "";// DNTC科室（主担）
						data3[(i * 10)+j][22] = "";// 设计担当
						data3[(i * 10)+j][23] = "";// 对策
						data3[(i * 10)+j][24] = "";// 对策导入时间
						data3[(i * 10)+j][25] = "";// 对策图片
						data3[(i * 10)+j][26] = "";// 结案判断
						data3[(i * 10)+j][27] = str[5];// 生产要件编号
						data3[(i * 10)+j][28] = "";// 分类（1）
						data3[(i * 10)+j][29] = "";// 分类（2）
						data3[(i * 10)+j][30] = "";// 版次
						data3[(i * 10)+j][31] = "";// 判断
						data3[(i * 10)+j][32] = "";// 判断要望
						data3[(i * 10)+j][33] = "";// 重大课题
					}
				}
				// 根据模板创建Excel模板
				XSSFWorkbook book = OutputDataToExcel3.creatXSSFWorkbook(inputStream, uplist, downlist, coverlist,commonlist);

				viewPanel.addInfomation("开始写数据，请耐心等待...\n", 70, 100);

				// 写数据
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

				// 删除空Sheet页
				if(!(uplist.isEmpty()&&downlist.isEmpty()&&coverlist.isEmpty()&&commonlist.isEmpty())) {
					OutputDataToExcel3.deleteXSSFWorkbook(book, uplist, downlist, coverlist, commonlist);
				}			
				// 输出打开文件
				String familycode = topbomline.getItemRevision().getProperty("project_ids");// 车型
				String vehicle = Util.getDFLProjectIdVehicle(familycode);
				if(vehicle==null || vehicle.isEmpty()) {
					vehicle = familycode;
				}
				String date = dateformat.format(new Date());
				// 输出的文件名称
				String datasetname = vehicle+"生产要件不具合一元表("+stage+")"+"_"+date+"时";
				String fileName = Util.formatString(datasetname);
				OutputDataToExcel3.exportFile(book, fileName.trim());
				//viewPanel.addInfomation("输出报表完成...\n", 100, 100);
				Util.saveFiles(fileName.trim(),datasetname, folder, session,"AI");

				viewPanel.addInfomation("输出报表完成，请在选择保存的文件夹下查看！...\n", 100, 100);
//			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// 排序
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
