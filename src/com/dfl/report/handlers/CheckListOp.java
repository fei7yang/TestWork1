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
	private String stage;// 阶段
	private TCSession session;
	TCComponentBOMLine topbomline;
	private String[] value;// 属性数组
	private String[] value1;// 要件库属性数组
	ArrayList valuelist = new ArrayList();// // 测试实例属性集合
	ArrayList valuelist1 = new ArrayList();// // 要件库属性集合
	int number = 0;// 数据行数
	String[][] data;
	private TCComponent folder;
	SimpleDateFormat dateformat = new SimpleDateFormat("yyyyMMdd  HH");// 设置日期格式
	SimpleDateFormat dateformat2 = new SimpleDateFormat("yyyy.MM.dd");// 设置日期格式
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
			// 查询并导出模板
			InputStream inputStream = FileUtil.getTemplateFile("DFL_Template_CheckList");
			System.out.println("inputStream=" + inputStream);
	        
			if (inputStream == null) {
				MessageBox.post("错误：没有找到要件检查表模板，请联系系统管理员在TC中添加模板(名称为：DFL_Template_CheckList)", "错误",
						MessageBox.INFORMATION);
				return ;
			}
			
			// 界面显示进度并输出执行步骤
			ReportViwePanel viewPanel = new ReportViwePanel("生成报表");
			viewPanel.setVisible(true);

			// 获取选择的对象
			//InterfaceAIFComponent[] aifComponents = app.getTargetComponents();
			TCComponentBOMLine bomline = (TCComponentBOMLine) aifComponents[0];
			topbomline = bomline.window().getTopBOMLine();

			viewPanel.addInfomation("开始输出报表...\n", 10, 100);
			

			// 先获取所有的虚层产线
			List lineList = Util.getChildrenByParent(aifComponents);
			// 获取所有焊装工位工艺关联的测试用例
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
					// 获取测试用例实例
					TCComponent testCaseInstance = modelObject.getTestComponent();
					// 获取测试用例
					TCComponent testCase = modelObject.getTestCase();
					// 获取测试用例的类型
					String testcasetype = Util.getRelProperty(testCase, "b8_TestCaseType");
					// 获取生产要件（类型为0）
					if (testcasetype.equals("0")) {
						if ((testCase != null && testCase instanceof TCComponentItem)
								&& (testCaseInstance != null && testCaseInstance instanceof TCComponentForm)) {

							// 获取测试用例最新版本
							TCComponentItem testCaseItem = (TCComponentItem) testCase;
							TCComponentItemRevision testCaseRev = testCaseItem.getLatestItemRevision();
							String distinguish = Util.getProperty(testCaseRev, "b8_distinguish");// 区分
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
							// 获取测试实例活动对象表单（测试活动）
							TCComponent[] activitys = testCaseInstance
									.getRelatedComponents("Tm0TestInstanceActivityRel");
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

									if (totalmap.containsKey(distinguish)) {
										Integer[] numbers = totalmap.get(distinguish);
										numbers[1] = numbers[1] + 1;
										totalmap.put(distinguish, numbers);
									}

									// 需要排序（仅取最新的测试活动）
									// //////////////////按活动日期进行排序
									Collections.sort(tempList2, comparator);

									// 取tempList中第一个为最新日期的测试活动
									TCComponentForm testactivity = (TCComponentForm) tempList2.get(0);
									// 开始获取表头属性
									value = new String[14];
									// 查找要件库中，测试实例的父级对象，获取其属性
									TCComponent[] relcomps = testCaseItem.whereUsed(TCComponent.WHERE_USED_ALL);
									if (relcomps != null && relcomps.length > 0) {
										
										int relcompcount = 0;

										for (int j = 0; j < relcomps.length; j++) {

											TCComponentItemRevision relcmp = (TCComponentItemRevision) relcomps[j];

											if (relcmp.isTypeOf("B8_RequirementRevision")) {
												// 获取表头第2列和第8列的属性
												value1 = new String[2];
												value1[0] = Util.getProperty(relcmp, "object_name");// 部位别名称：需在生产要件上新增属性
												value1[1] = distinguish;// 区分
												valuelist1.add(value1);

												value[0] = "";// 模板表头前空一格
												value[2] = Util.getProperty(testCaseRev, "b8_TestCaseID");// 生产要件编号
												// 获取部位别编号，去零
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

												value[1] = no;// 部位别编号：生产要件编号第一个“-”前的字段，首位0则去掉0

												value[3] = Util.getProperty(testCaseRev, "object_name");// 生产要件名称
												value[4] = Util.getProperty(testCaseRev, "b8_Level");// MUST&WANT区分（级别）
												value[5] = "";// 采用（默认为空）
												value[6] = Util.getProperty(testCaseRev, "b8_PhaseIn");// 纳入时期
												String status = Util.getProperty(testactivity, "tm0ResultStatus");// 确认结果:新增此属性，采用系统已有的测试结果LOV,根据表头填充颜色
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

												value[8] = Util.getProperty(testactivity, "tm0Comment");// 判定内容
												value[9] = Util.getUserName(testactivity);// 确认人
												value[10] = Util.getProperty(testactivity, "tm0ActivityDate");// 确认日期(活动日期)
												value[11] = "";// 确认结果
												value[12] = "";// 确认日期
												value[13] = "";// 备注
												valuelist.add(value);
												number++;
												
												relcompcount++ ;
											}
										}
									} else {
										value1 = new String[2];
										value1[0] = "";// 部位别名称：需在生产要件上新增属性
										value1[1] = distinguish;// 区分
										valuelist1.add(value1);
										value[0] = "";// 模板表头前空一格
										value[2] = Util.getProperty(testCaseRev, "b8_TestCaseID");// 生产要件编号
										// 获取部位别编号，去零
										String val = value[2];
										String no = "";
										if (val.contains("-")) {
											no = val.substring(0, val.indexOf("-"));
											if (no.startsWith("0") && no.length() > 1) {
												no = no.substring(1, no.length());
											}
										}
										value[1] = no;// 部位别编号：生产要件编号第一个“-”前的字段，首位0则去掉0

										value[3] = Util.getProperty(testCaseRev, "object_name");// 生产要件名称
										value[4] = Util.getProperty(testCaseRev, "b8_Level");// MUST&WANT区分（级别）
										value[5] = "";// 采用（默认为空）
										value[6] = Util.getProperty(testCaseRev, "b8_PhaseIn");// 纳入时期
										String status = Util.getProperty(testactivity, "tm0ResultStatus");// 确认结果:新增此属性，采用系统已有的测试结果LOV,根据表头填充颜色
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
										value[8] = Util.getProperty(testactivity, "tm0Comment");// 判定内容
										value[9] = Util.getUserName(testactivity);// 确认人
										value[10] = Util.getProperty(testactivity, "tm0ActivityDate");// 确认日期(活动日期)
										value[11] = "";// 确认结果
										value[12] = "";// 确认日期
										value[13] = "";// 备注
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

				viewPanel.addInfomation("正在输出报表...\n", 35, 100);
				// 定义报表行数据
				data = new String[number][16];
				ArrayList datalist = new ArrayList();
				for (int i1 = 0; i1 < valuelist.size(); i1++) {
					String[] values = new String[17];
					String[] str1 = (String[]) valuelist1.get(i1);
					String[] str = (String[]) valuelist.get(i1);

					values[0] = str[0];// 模板表头前空一格
					values[1] = str[1];// 部位别编号
					values[2] = str1[0];// 部位别名称：需在生产要件上新增属性
					values[3] = str[2];// 生产要件编号
					values[4] = str[3];// 生产要件名称
					values[5] = str[4];// MUST&WANT区分（级别）
					values[6] = str[5];// 采用
					values[7] = str[6];//// 纳入时期
					values[8] = str1[1];// 区分
					values[9] = str[7];// 确认结果:新增此属性，根据表头填充颜色
					values[10] = str[8];// 判定内容
					values[11] = str[9];// 确认人
					values[12] = str[10];// 确认日期
					values[13] = "";
					values[14] = str[11];// 确认结果
					values[15] = str[12];// 确认日期
					values[16] = str[13];// 备注
					datalist.add(values);
				}
				// 根据部位别编号排序
				Comparator comparator2 = getComParatorBySerialID();
				Collections.sort(datalist, comparator2);
				String familycode = topbomline.getItemRevision().getProperty("project_ids");// 车型
				String vehicle = Util.getDFLProjectIdVehicle(familycode);
				if(vehicle==null || vehicle.isEmpty()) {
					vehicle = familycode;
				}
				//去掉COMMON
				if(totalmap.containsKey("COMMON")) {
					totalmap.remove("COMMON");
				}
				ArrayList totallist = new ArrayList();
				totallist.add(totalmap);
				totallist.add(vehicle);
				totallist.add(dateformat2.format(new Date()));
				totallist.add(stage);
				
				viewPanel.addInfomation("开始写数据，请耐心等待...\n", 70, 100);
				// 输出的文件名称

				String date = dateformat.format(new Date());
				String datasetname = vehicle + "生产要件检查表(" + stage + ")" + "_" + date + "时";
				String fileName = Util.formatString(datasetname);
				OutputDataToExcel1 outdata = new OutputDataToExcel1(datalist, totallist, inputStream, fileName.trim());
				// 输出文件
				// viewPanel.addInfomation("输出报表完成...\n", 100, 100);
				Util.saveFiles(fileName.trim(), datasetname, folder, session, "AP");

				viewPanel.addInfomation("输出报表完成，请在选择保存的文件夹下查看！...\n", 100, 100);
			
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
