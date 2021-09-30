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
	private String stage;// 阶段
	private TCComponent folder;
	private InterfaceAIFComponent[] aifComponents;

	SimpleDateFormat dateformat = new SimpleDateFormat("yyyyMMdd  HH");// 设置日期格式

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
			ArrayList datalist = new ArrayList();// 所有数据集合
			// 界面显示进度并输出执行步骤
			ReportViwePanel viewPanel = new ReportViwePanel("生成报表");
			viewPanel.setVisible(true);

			// 获取选择的对象
			//InterfaceAIFComponent[] aifComponents = app.getTargetComponents();
			TCComponentBOMLine bomline = (TCComponentBOMLine) aifComponents[0];
			topbomline = bomline.window().getTopBOMLine();

			viewPanel.addInfomation("正在获取模板...\n", 30, 100);
			// 查询并导出模板
			InputStream inputStream = FileUtil.getTemplateFile("DFL_Template_WatertightCheckList");

			if (inputStream == null) {

				/*
				 * viewPanel.addInfomation(
				 * "错误：没有找到水密要件检查表模板，请先在TC中添加模板(名称为：DFL_Template_WatertightCheckList)\n", 100,
				 * 100);
				 */
				System.out.println("错误：没有找到水密要件检查表模板，请先在TC中添加模板(名称为：DFL_Template_WatertightCheckList)");
				return;
			}

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

			// 水密要件位置图片，关系B8_TestCase_Watertight_Pos下
			ArrayList pngmap = new ArrayList();

			if (list != null && list.size() > 0) {
				for (int i = 0; i < list.size(); i++) {
					TestManagerModelObject modelObject = (TestManagerModelObject) list.get(i);
					// 获取测试用例实例
					TCComponent testCaseInstance = modelObject.getTestComponent();
					// 获取测试用例
					TCComponent testCase = modelObject.getTestCase();
					// 获取类型为1的测试用例（水密要件）
					String testcasetype = Util.getRelProperty(testCase, "b8_TestCaseType");
					if (testcasetype.equals("1")) {
						if ((testCase != null && testCase instanceof TCComponentItem)
								&& (testCaseInstance != null && testCaseInstance instanceof TCComponentForm)) {
							// 获取测试用例最新版本
							TCComponentItem testCaseItem = (TCComponentItem) testCase;
							TCComponentItemRevision testCaseRev = testCaseItem.getLatestItemRevision();

							// 获取测试实例活动对象表单（测试活动）
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

									// 取tempList中测试活动
									TCComponentForm testactivity = (TCComponentForm) tempList2.get(0);
									// 开始获取表头属性
									value = new Object[13];
									value[0] = Util.getProperty(testCaseRev, "b8_SerialID");// 序号NO，需进行排序
									value[1] = Util.getProperty(testCaseRev, "b8_distinguish");// 区域
									value[2] = Util.getProperty(testCaseRev, "b8_TestCasePart");// 部位
									value[3] = Util.getProperty(testCaseRev, "b8_ApplicableCar");// 适用构造
									value[4] = Util.getProperty(testCaseRev, "b8_DefectReason");// 不具合要因
									// 水密需要注意部位，图片挂载关系B8_TestCase_Watertight
									ArrayList phomap = new ArrayList();
									// 获取B8_TestCase_Watertight关系下的数据集对象
									TCComponent[] tdata = testCaseRev.getRelatedComponents("B8_TestCase_Watertight");
									// 将图片类型的数据集下载到本地
									for (TCComponent tdt : tdata) {
										File file = Util.downLoadPicture(tdt);
										if (file != null) {
											phomap.add(file);
										}
									}

									value[5] = phomap;// 水密需要注意部位图片
									value[6] = Util.getProperty(testCaseRev, "b8_Check");// 检查项目
									value[7] = Util.getProperty(testCaseRev, "b8_Remarks");// 备注
									value[8] = Util.getProperty(testactivity, "b8_TestStage");// 阶段
									String status = Util.getProperty(testactivity, "tm0ResultStatus");// 採否（对应测试结果）
									value[9] = status.substring(0, 1);// 对应测试结果前一位

									// 获取测试结果关系下的数据集对象(现状图片)
									ArrayList picmap = new ArrayList();
									TCComponent[] tdata1 = testactivity.getRelatedComponents("Tm0TestResultRel");
									// 将图片类型的数据集下载到本地
									for (TCComponent tdt1 : tdata1) {
										File file = Util.downLoadPicture(tdt1);
										if (file != null) {
											picmap.add(file);
										}
									}
									value[10] = picmap;// 现状图示

									value[11] = Util.getProperty(testactivity, "tm0Comment");// 问题描述

									TCComponent[] tdata2 = testCaseRev
											.getRelatedComponents("B8_TestCase_Watertight_Pos");
									// 将图片类型的数据集下载到本地
									for (TCComponent tdt2 : tdata2) {
										File file = Util.downLoadPicture(tdt2);
										if (file != null) {
											if (pngmap.size() < 1) {
												pngmap.add(file);
											}
										}
									}
									value[12] = pngmap;// 水密要件位置图片

									datalist.add(value);
								}
							}
						}
					}
				}
			}
			viewPanel.addInfomation("开始写数据，请耐心等待...\n", 50, 100);

			Comparator comparator2 = getComParatorBySerialID();
			Collections.sort(datalist, comparator2);
			// 根据模板创建Excel模板
			XSSFWorkbook book = OutputDataToExcel3.creatXSSFWorkbook(inputStream);

			viewPanel.addInfomation("", 80, 100);
			// 写数据
			OutputDataToExcel3.writeDataToSheet(book, datalist);

			// 输出文件
			String familycode = topbomline.getItemRevision().getProperty("project_ids");// 车型
			String vehicle = Util.getDFLProjectIdVehicle(familycode);
			if(vehicle==null || vehicle.isEmpty()) {
				vehicle = familycode;
			}
			String date = dateformat.format(new Date());
			// 输出的文件名称
			String datasetname = vehicle + "水密要件检查结果一元表(" + stage + ")" + "_" + date + "时";
			String fileName = Util.formatString(datasetname);
			OutputDataToExcel3.exportFile(book, fileName.trim());
			// viewPanel.addInfomation("输出报表完成...\n", 100, 100);
			Util.saveFiles(fileName.trim(), datasetname, folder, session, "AQ");

			viewPanel.addInfomation("输出报表完成，请在选择保存的文件夹下查看！...\n", 100, 100);
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
