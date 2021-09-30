package com.dfl.report.handlers;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.Dictionary;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import javax.swing.SwingUtilities;

import com.dfl.report.util.FileUtil;
import com.dfl.report.util.OutputDataToExcel;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.kernel.*;
import com.teamcenter.rac.pse.AbstractPSEApplication;
import com.teamcenter.rac.util.MessageBox;

public class RobotAndWeldingExportOp {

	public RobotAndWeldingExportOp(AbstractAIFUIApplication app, InputStream inputStream) {
		// TODO Auto-generated constructor stub
		this.app = app;
		this.session = (TCSession) this.app.getSession();
		this.inputStream = inputStream;
		initUI();
	}

	private AbstractAIFUIApplication app;
	private TCSession session;
	ArrayList<TCComponentBOMLine> meresource = new ArrayList<TCComponentBOMLine>();// 第一层资源曾集合
	private String[] medata;// 数据集合
	private ArrayList data = new ArrayList();
	private InputStream inputStream;
    private String error = "";

	private void initUI() {
		// TODO Auto-generated method stub
		try {
			// 界面显示进度并输出执行步骤
			ReportViwePanel viewPanel = new ReportViwePanel("生成报表");
			viewPanel.setVisible(true);

			// 获取选择的对象
			InterfaceAIFComponent[] ifc = app.getTargetComponents();

			TCComponentBOMLine topbomline = (TCComponentBOMLine) ifc[0];

			String linename = Util.getProperty(topbomline, "bl_rev_object_name");

			// viewPanel.addInfomation("开始输出报表...\n", 5, 100);

			// 判断用户对所选对象是否有写权限
//			boolean flag = Util.hasWritePrivilege(session, topbomline.getItemRevision());
//			if (!flag) {
//				viewPanel.addInfomation("对当前焊装产线没有写权限！...\n", 100, 100);
//				return;
//			}
			viewPanel.addInfomation("正在获取模板...\n", 10, 100);
			// 查询导出模板
//			InputStream inputStream = FileUtil.getTemplateFile("DFL_Template_RobotAndWeldingList");
//
//			if (inputStream == null) {
//				viewPanel.addInfomation("错误：没有找到机器人&焊枪清单模板，请先添加模板(名称为：DFL_Template_RobotAndWeldingList)\n", 100, 100);
//				return;
//			}

			// 机器人和焊枪B8_BIWRobotRevision和B8_BIWGunRevision
			long startTime = System.currentTimeMillis(); // 获取开始时间

			viewPanel.addInfomation("开始输出报表...\n", 20, 100);

			// 获取第一资源层集合
			getStationBomLine(topbomline, viewPanel);

			long endTime = System.currentTimeMillis(); // 获取结束时间
			System.out.println("获取机器人和焊枪程序运行时间： " + (endTime - startTime) + "ms");

			// 循环焊装工位获取下层的机器人或焊枪一共有多少层
			startTime = System.currentTimeMillis(); // 获取开始时间

			viewPanel.addInfomation("", 40, 100);

			if (meresource.size() > 0) {
				data = getRobotGunPropertys(meresource, viewPanel, 40);// 数据集合
			}
			if(!error.isEmpty())
			{
				viewPanel.dispose();
				inputStream.close();
				MessageBox.post("以下工位的资源包中存在多个机器人，请检查数据：" + error, "温馨提示", MessageBox.INFORMATION);
				return;
			}
			
			endTime = System.currentTimeMillis(); // 获取结束时间
			System.out.println("获取属性程序运行时间： " + (endTime - startTime) + "ms");

			viewPanel.addInfomation("开始写数据，请耐心等待...\n", 60, 100);

//			Comparator comparator2 = getComParator();
//			Collections.sort(data, comparator2);

			viewPanel.addInfomation("", 80, 100);
			String datasetname = linename + "机器人&焊枪清单";
			String filename = Util.formatString(datasetname);

			OutputDataToExcel odte = new OutputDataToExcel(data, inputStream, filename, viewPanel);

			saveFiles(datasetname, filename, topbomline);

			viewPanel.addInfomation("输出报表完成，请在焊装产线版本附件下查看报表\n", 100, 100);

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// 用递归算法获取第一资源层MECompResourceRevision
	public void getStationBomLine(TCComponentBOMLine parent, ReportViwePanel viewPanel) {

		viewPanel.addInfomation("", 10, 100);
		try {
			AIFComponentContext[] children = parent.getChildren();
			// String parentName = parent.getProperty("bl_rev_object_name");
			for (AIFComponentContext chld : children) {
				TCComponentItemRevision rev = ((TCComponentBOMLine) chld.getComponent()).getItemRevision();
				if (rev.isTypeOf("MECompResourceRevision")) {
					meresource.add((TCComponentBOMLine) chld.getComponent());
					continue;
				} else {
					getStationBomLine((TCComponentBOMLine) chld.getComponent(), viewPanel);
				}
			}
		} catch (TCException e) {
			// TODO 自动生成的 catch 块
			e.printStackTrace();
		}
	}

	// 循环机器人，获取机器人或焊枪相关属性值
	public ArrayList getRobotGunPropertys(ArrayList<TCComponentBOMLine> parent, ReportViwePanel viewPanel, int n) {
		ArrayList list = new ArrayList();

		// 根据顶层BOP查询所有的焊装产线
		String typename = Util.getObjectDisplayName(session, "B8_BIWRobot");
		String[] propertys = new String[] { "bl_item_object_type", "bl_item_object_type" };
		String[] values = new String[] { typename, typename };

		// 根据焊装产线查询所有的点焊工序
		String guntypename = Util.getObjectDisplayName(session, "B8_BIWGun");
		String[] propertys2 = new String[] { "bl_item_object_type", "bl_item_object_type" };
		String[] values2 = new String[] { guntypename, guntypename };
		int rownum = 1;
		for (int i = 0; i < parent.size(); i++) {
			viewPanel.addInfomation("", n, 100);
			TCComponentBOMLine bl = parent.get(i);
			ArrayList<TCComponentBOMLine> Robotbl = Util.searchBOMLine(bl, "OR", propertys, "==", values);// 机器人集合
			ArrayList<TCComponentBOMLine> Gunbl = Util.searchBOMLine(bl, "OR", propertys2, "==", values2);// 焊枪集合

			if(Robotbl!=null && Robotbl.size()>1)
			{
				String stateresource = "工位名称：" + getStationProperty(parent.get(i)) + " 资源包：" + Util.getProperty(parent.get(i), "bl_rev_object_name");
				if(error.isEmpty())
				{
					error = stateresource;
				}
				else
				{
					
					error = error + "," + stateresource;
									
				}			
			}
			
			// 根据焊枪集合确认输出几行数据
			if (Gunbl.size() > 0) {
				for (int j = 0; j < Gunbl.size(); j++) {
					medata = new String[14];
					medata[0] = Integer.toString(rownum); // 序号
					medata[1] = getStationProperty(parent.get(i));// 获取机器人对应的焊装工位
					medata[2] = Util.getProperty(parent.get(i), "bl_rev_object_name");// 机器人编号

					if (Robotbl.size() > 0) {
						TCComponentItemRevision rev;
						try {
							rev = Robotbl.get(0).getItemRevision();
							medata[3] = Util.getProperty(rev, "b8_Model");// 机器人型号
							medata[4] = Util.getProperty(rev, "b8_Features");// 机器人功能
						} catch (TCException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}

					} else {
						medata[3] = "";// 机器人型号
						medata[4] = "";// 机器人功能
					}
					medata[5] = getRobotHight(parent.get(i));// 机器人底座高度

					TCComponentItemRevision gunrev;
					try {
						gunrev = Gunbl.get(j).getItemRevision();
						medata[6] = Util.getProperty(gunrev, "b8_Model");// 焊枪枪号
						medata[7] = Util.getProperty(gunrev, "b8_MotorBrand");// 焊枪马达
						medata[8] = Util.getProperty(gunrev, "b8_Deep");// 焊枪喉深
						medata[9] = Util.getProperty(gunrev, "b8_Width");// 焊枪喉宽
						medata[10] = Util.getProperty(gunrev, "b8_Stroke");// 焊枪开口行程
						medata[11] = Util.getProperty(gunrev, "b8_MaxAmp");// 焊枪电流
						medata[12] = Util.getProperty(gunrev, "b8_Voltage1");// 焊枪电压
					} catch (TCException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					medata[13] = "";// 备注
					rownum++;
					list.add(medata);
				}

			} else {
				if (Robotbl.size() > 0) {
					medata = new String[14];
					medata[0] = Integer.toString(rownum); // 序号
					medata[1] = getStationProperty(parent.get(i));// 获取机器人对应的焊装工位
					medata[2] = Util.getProperty(parent.get(i), "bl_rev_object_name");// 机器人编号

					if (Robotbl.size() > 0) {
						TCComponentItemRevision rev;
						try {
							rev = Robotbl.get(0).getItemRevision();
							medata[3] = Util.getProperty(rev, "b8_Model");// 机器人型号
							medata[4] = Util.getProperty(rev, "b8_Features");// 机器人功能
						} catch (TCException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}

					} else {
						medata[3] = "";// 机器人型号
						medata[4] = "";// 机器人功能
					}
					medata[5] = getRobotHight(parent.get(i));// 机器人底座高度

					medata[6] = "";// 焊枪枪号
					medata[7] = "";// 焊枪马达
					medata[8] = "";// 焊枪喉深
					medata[9] = "";// 焊枪喉宽
					medata[10] = "";// 焊枪开口行程
					medata[11] = "";// 焊枪电流
					medata[12] = "";// 焊枪电压
					medata[13] = "";// 备注

					rownum++;
					list.add(medata);
				}

			}
		}
		return list;
	}

	// 根据机器人获取对应的焊装工位信息
	public String getStationProperty(TCComponentBOMLine chilren) {
		String station_name = "";
		try {
			TCComponentBOMLine station_bl = chilren.parent();
			TCComponentItemRevision rev = station_bl.getItemRevision();
			if (rev.isTypeOf("B8_StationRevision")) {
				station_name = Util.getProperty(station_bl, "bl_rev_object_name");
//				System.out.println(station_name);
				return station_name;
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
		return station_name;
	}

	// 根据机器人获取对应的机器人底座高度
	public String getRobotHight(TCComponentBOMLine chilren) {
		String station_name = "";
		try {
			// 第一层复合资源层
			AIFComponentContext[] chilrens = chilren.getChildren();
			for (AIFComponentContext chid : chilrens) {
				AIFComponentContext[] minchilrens = ((TCComponentBOMLine) chid.getComponent()).getChildren();// 第二层复合资源层
				for (AIFComponentContext ch : minchilrens) {
					TCComponentItemRevision rev = ((TCComponentBOMLine) ch.getComponent()).getItemRevision();
					if (rev.isTypeOf("B8_BIWDeviceRevision")) {
						station_name = Util.getProperty(rev, "b8_Spec");
						continue;
					}
				}

			}
			return station_name;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return station_name;
	}

	private Comparator getComParator() {
		// TODO Auto-generated method stub
		Comparator comparator = new Comparator() {

			public int compare(Object obj, Object obj1) {
				Object[] comp1 = (Object[]) obj;
				Object[] comp2 = (Object[]) obj1;

				return comp1[1].toString().compareTo(comp2[1].toString());
			}
		};

		return comparator;
	}

	// 把生成的报表，作为数据集放到Newstaff文件夹下
	public void saveFiles(String datasetname, String filename, TCComponentBOMLine topbomline) {
		try {
			String fullFileName = FileUtil.getReportFileName(filename);
			TCSession session = (TCSession) app.getSession();
			TCComponentUser user = session.getUser();
			//TCComponentFolder newstuff = user.getNewStuffFolder();

			TCComponentItemRevision rev = topbomline.getItemRevision();
			TCComponentItemRevision docrev = null;
			// 判断是否之前已生成过文档
			TCComponent[] children = rev.getRelatedComponents("IMAN_reference");
			if (children.length > 0) {
				for (TCComponent child : children) {
					if (child instanceof TCComponentItem) {
						TCComponentItem item = (TCComponentItem) child;
						if (item.isTypeOf("B8_BIWProcDoc")) {
							docrev = item.getLatestItemRevision();
							break;
						}
					}
				}
			}
			TCComponentDataset ds = Util.createDataset(session, datasetname, fullFileName, "MSExcelX", "excel");
			if (docrev == null) {
				TCComponentItemType tcccomponentitemtype = (TCComponentItemType) session
						.getTypeComponent("B8_BIWProcDoc");

				TCComponentItem tccomponentitem = tcccomponentitemtype.create("", "", "B8_BIWProcDoc", datasetname,
						"desc", null);
				tccomponentitem.setProperty("b8_BIWProcDocType", "AH");
				tccomponentitem.lock();
				tccomponentitem.save();
				tccomponentitem.unlock();
				docrev = tccomponentitem.getLatestItemRevision();
				
				// 添加产线版本与文档的关系
				rev.add("IMAN_reference", docrev.getItem());
				rev.lock();
				rev.save();
				rev.unlock();
			} else {
				// 先移除之前生成的报表
				TCComponent[] childrens = docrev.getRelatedComponents("IMAN_specification");
				if (children.length > 0) {
					for (TCComponent child : childrens) {
						if (child instanceof TCComponentDataset) {
							TCComponentDataset dataset = (TCComponentDataset) child;
							docrev.cutOperation("IMAN_specification", new TCComponent[] { dataset });
							dataset.delete();
						}
					}
				}
			}
			// 添加文档与数据集的关系
			docrev.add("IMAN_specification", ds);
			docrev.lock();
			docrev.save();
			docrev.unlock();
			

			// 删除中间文件
			File file = new File(fullFileName);
			if (file.isFile()) {
				file.delete();
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
