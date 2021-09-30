package com.dfl.report.handlers;

import java.awt.Container;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.swt.widgets.DateTime;

import com.dfl.report.common.TCComponentUtils;
import com.dfl.report.util.FileUtil;
import com.dfl.report.util.NewOutputDataToExcel;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentDatasetType;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentItemType;
import com.teamcenter.rac.kernel.TCComponentUser;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class HitTheTableExportOp {

	private AbstractAIFUIApplication app;
	ArrayList<TCComponentBOMLine> list = new ArrayList<TCComponentBOMLine>();// 点焊工序集合
	private String[] Discrete;// 点焊信息数据 {点焊工序名称，日期，机器人编号，机器人型号，焊枪枪号，焊装工位，车型}
	private String[] Weld;// 焊点信息数组{点焊工序名称，焊点ID，序号}
	ArrayList Discretelist = new ArrayList();// 点焊信息集合
	ArrayList Weldlist = new ArrayList();// 焊点信息集合
	String vehicle = null; // 车型
	private String beginname;
	private InputStream inputStream;
	private boolean Isupdateflag;
	private TCComponentItem tcc = null;

	public HitTheTableExportOp(AbstractAIFUIApplication app, InputStream inputStream, boolean isupdateflag, TCComponentItem tcc) {
		// TODO Auto-generated constructor stub
		this.app = app;
		this.inputStream = inputStream;
		this.Isupdateflag = isupdateflag;
		this.tcc = tcc;
		initUI();
	}

	private void initUI() {
		// TODO Auto-generated method stub
		try {

			// 获取选择的对象
			InterfaceAIFComponent[] ifc = app.getTargetComponents();

			TCComponentBOMLine topbomline = (TCComponentBOMLine) ifc[0];

			TCSession session = (TCSession) app.getSession();
			TCComponentUser user = session.getUser();
			TCComponentFolder newstuff = user.getNewStuffFolder();
			TCComponent[] projects = topbomline.window().getTopBOMLine().getItemRevision().getRelatedComponents("project_list");

			// 根据选择的焊装产线工艺下，是否已经生成过报表，如果生成过则直接取之前的报表作为模板
			TCComponentItemRevision blrev = topbomline.getItemRevision();

			// 输出的文件名称
			String datasetname = topbomline.getProperty("bl_rev_object_name") + "打顺表";
			//String fileName = Util.formatString(datasetname);
			String fileName = datasetname;
			
				
			if (tcc == null) {
				// 获取topbomline上一级
				TCComponentBOMLine uplevel = topbomline.window().getTopBOMLine();

				TCComponentItemRevision uprev = uplevel.getItemRevision();

				// vehicle = Util.getProperty(uplevel, "b8_vehicle_id");
				vehicle = Util.getProperty(uprev, "project_ids");

				// 界面显示进度并输出执行步骤
				ReportViwePanel viewPanel = new ReportViwePanel("生成报表");
				viewPanel.setVisible(true);

				// viewPanel.addInfomation("开始输出报表...\n", 5, 100);
				
				viewPanel.addInfomation("开始输出报表...\n", 10, 100);
				
				viewPanel.addInfomation("正在输出报表...\n", 20, 100);

				// 用递归算法获取所有的点焊工序B8_BIWDiscreteOP
				getBIWDiscreteOP(topbomline);

				viewPanel.addInfomation("", 40, 100);

				// 判断是否有点焊工序数据
				if (list.size() < 1) {
					inputStream.close();
					viewPanel.dispose();
					MessageBox.post("错误：当前焊装产线工艺没有机器人点焊工序数据", "温馨提示", MessageBox.INFORMATION);
					//viewPanel.addInfomation("错误：当前焊装产线工艺没有点焊工序数据\n", 100, 100);
					return;
				}

				// 循环开始点焊工序名称标记
				beginname = ((TCComponentBOMLine) list.get(0)).getProperty("bl_rev_object_name");

				// 循环点焊工序获取机器人、焊点等信息
				getStationBomLine(list);

				// 判断是否有点焊工序数据
				if (Discretelist.size() < 1 && Weldlist.size() < 1) {
					inputStream.close();
					viewPanel.dispose();
					MessageBox.post("错误：当前焊装产线工艺没有机器人点焊工序数据", "温馨提示", MessageBox.INFORMATION);
					//viewPanel.addInfomation("错误：当前焊装产线工艺没有机器人和焊点数据\n", 100, 100);
					return;
				}

				viewPanel.addInfomation("开始写数据，请耐心等待...\n", 60, 100);
				// 写excel数据
				// 根据模板创建Excel空模板
				ArrayList errorname = new ArrayList();
				XSSFWorkbook book = NewOutputDataToExcel.creatXSSFWorkbook(inputStream, Discretelist,errorname);

				if (book == null) {
					viewPanel.dispose();
					String statename = "";
					for(int i=0;i<errorname.size();i++)
					{
						if(i == 0)
						{
							statename = (String) errorname.get(i);
						}
						else
						{
							statename = statename + "、" + errorname.get(i);
						}
					}
					MessageBox.post("错误：当前焊装产线工艺存在:" + statename + "的sheet命名长度过长，无法生成打顺表", "温馨提示", MessageBox.INFORMATION);
					return;
				}

				// 写sheet数据，点焊工序数据写入
				NewOutputDataToExcel.writeDiscreteDataToSheet(book, Discretelist,Weldlist, true);

				viewPanel.addInfomation("", 80, 100);

				// 写sheet数据，焊点数据写入
				NewOutputDataToExcel.writeWeldDataToSheet(book, Weldlist,Discretelist, true);

				// 输出文件
				NewOutputDataToExcel.exportFile(book, fileName);

				saveFiles(datasetname, fileName, topbomline,tcc,projects);
								

				viewPanel.addInfomation("输出报表完成，请在焊装产线工艺对象附件下或Newstuff文件夹下查看报表\n", 100, 100);

			} else {
				// 下载文件到本地

				// 界面显示进度并输出执行步骤
				ReportViwePanel viewPanel = new ReportViwePanel("生成报表");
				viewPanel.setVisible(true);
	
				viewPanel.addInfomation("开始输出报表...\n", 10, 100);
	
				// 获取topbomline上一级
				TCComponentBOMLine uplevel = topbomline.window().getTopBOMLine();
				// 从BOP版本对象上去车型信息
				TCComponentItemRevision uprev = uplevel.getItemRevision();

				String tempvehicle = Util.getProperty(uprev, "project_ids");
				
				vehicle = Util.getDFLProjectIdVehicle(tempvehicle);
				
				if(vehicle==null || vehicle.isEmpty()) {
					vehicle = tempvehicle;
				}

				// 用递归算法获取所有的点焊工序B8_BIWDiscreteOP
				getBIWDiscreteOP(topbomline);

				viewPanel.addInfomation("正在输出报表...\n", 20, 100);

				// 判断是否有点焊工序数据
				if (list.size() < 1) {
					inputStream.close();
					viewPanel.dispose();
					MessageBox.post("错误：当前焊装产线工艺没有机器人点焊工序数据", "温馨提示", MessageBox.INFORMATION);
					//viewPanel.addInfomation("错误：当前焊装产线工艺没有点焊工序数据\n", 100, 100);
					return;
				}

				// 循环开始点焊工序名称标记
				beginname = ((TCComponentBOMLine) list.get(0)).getProperty("bl_rev_object_name");

				// 循环点焊工序获取机器人、焊点等信息
				getStationBomLine(list);

				// 判断是否有点焊工序数据
				if (Discretelist.size() < 1 && Weldlist.size() < 1) {
					inputStream.close();
					viewPanel.dispose();
					MessageBox.post("错误：当前焊装产线工艺没有机器人点焊工序数据", "温馨提示", MessageBox.INFORMATION);
					//viewPanel.addInfomation("错误：当前焊装产线工艺没有机器人和焊点数据\n", 100, 100);
					return;
				}
				viewPanel.addInfomation("开始写数据，请耐心等待...\n", 40, 100);

				XSSFWorkbook book = null;
				if(Isupdateflag) {
					System.out.println("进入更新方法 " + Isupdateflag);
					ArrayList errorname = new ArrayList();
					// 根据模板创建Excel模板
					book = NewOutputDataToExcel.updateXSSFWorkbook(inputStream, Discretelist,errorname);
					
					if (book == null) {
						viewPanel.dispose();
						String statename = "";
						for(int i=0;i<errorname.size();i++)
						{
							if(i == 0)
							{
								statename = (String) errorname.get(i);
							}
							else
							{
								statename = statename + "、" + errorname.get(i);
							}
						}
						MessageBox.post("错误：当前焊装产线工艺存在:" + statename + "的sheet命名长度过长，无法生成打顺表", "温馨提示", MessageBox.INFORMATION);
					}
					
					// 写sheet数据，点焊工序数据写入
					NewOutputDataToExcel.UpdateDiscreteDataToSheet(book, Discretelist, Weldlist,true);
					viewPanel.addInfomation("", 60, 100);
					// 更新sheet数据，焊点数据写入
					NewOutputDataToExcel.updateWeldDataToSheet(book, Weldlist, Discretelist, true);
				}else {
					System.out.println("进入新增方法 " + Isupdateflag);
					ArrayList errorname = new ArrayList();
					book = NewOutputDataToExcel.creatXSSFWorkbook(inputStream, Discretelist,errorname);
					
					if (book == null) {
						viewPanel.dispose();
						String statename = "";
						for(int i=0;i<errorname.size();i++)
						{
							if(i == 0)
							{
								statename = (String) errorname.get(i);
							}
							else
							{
								statename = statename + "、" + errorname.get(i);
							}
							
						}
						MessageBox.post("错误：当前焊装产线工艺存在:" + statename + "的sheet命名长度过长，无法生成打顺表", "温馨提示", MessageBox.INFORMATION);
						return;
					}
					
					// 写sheet数据，点焊工序数据写入
					NewOutputDataToExcel.writeDiscreteDataToSheet(book, Discretelist,Weldlist, true);
					viewPanel.addInfomation("", 60, 100);
					// 更新sheet数据，焊点数据写入
					NewOutputDataToExcel.writeWeldDataToSheet(book, Weldlist,Discretelist, true);
				}
				
				viewPanel.addInfomation("", 80, 100);
				// 输出文件
				NewOutputDataToExcel.exportFile(book, fileName);

				saveFiles(datasetname, fileName, topbomline,tcc,projects);
				
				// 把更新的报表上传
//				String filePath = FileUtil.getReportFileName(fileName);
//				String nameRef = "excel";
//				String[] filePaths = { filePath };
//				String[] namedRefs = { nameRef };
//				dataset.setFiles(filePaths, namedRefs);
//				dataset.lock();
//				dataset.save();
//				dataset.unlock();

				viewPanel.addInfomation("输出报表完成，请在焊装产线工艺对象附件下或Newstuff文件夹下查看报表\n", 100, 100);
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// 把生成的报表，作为数据集放到Newstaff文件夹下
	public void saveFiles(String datasetname, String filename, TCComponentBOMLine topbomline, TCComponentItem tcc2, TCComponent[] projects) {
		try {
			TCComponentItemRevision toprev = topbomline.getItemRevision();
			TCSession session = (TCSession) app.getSession();
			TCComponentItem tccomponentitem;
			String fullFileName = FileUtil.getReportFileName(filename);
			TCComponentUser user = session.getUser();
			TCComponentFolder newstuff = user.getNewStuffFolder();
			TCComponentItemType tcccomponentitemtype = (TCComponentItemType) session.getTypeComponent("B8_BIWProcDoc");

			if(tcc == null) {
				tccomponentitem = tcccomponentitemtype.create("", "", "B8_BIWProcDoc", datasetname, "desc",
						null);
				newstuff.add("contents", tccomponentitem);
				// MessageBox.post("创建成功！", "信息提示", MessageBox.INFORMATION);
				tccomponentitem.setProperty("b8_BIWProcDocType", "AA");
				tccomponentitem.lock();
				tccomponentitem.save();
				tccomponentitem.unlock();
				// 添加焊装产线与文档的关系
				toprev.add("IMAN_reference", tccomponentitem);
				toprev.lock();
				toprev.save();
				toprev.unlock();
			}else {
				tccomponentitem = tcc;
			}						
			TCComponentItemRevision rev = tccomponentitem.getLatestItemRevision();
			// 移除的时候，需要将所有符合条件的都查找出来，再移除
			TCComponent[] children = TCComponentUtils.getCompsByRelation(rev, "IMAN_specification");
			for (TCComponent child : children) {
				if (child instanceof TCComponentDataset) {
					TCComponentDataset dataset = (TCComponentDataset) child;
					rev.cutOperation("IMAN_specification", new TCComponent[] { dataset });
					try {
						dataset.delete();
					} catch (Exception e2) {

					}
				}
			}					
			TCComponentDataset ds = Util.createDataset(session, datasetname, fullFileName, "MSExcelX", "excel");
			// 添加文档与数据集的关系
			rev.add("IMAN_specification", ds);
			rev.lock();
			rev.save();
			rev.unlock();
			
			// 将文档指派项目
			Util.assignProjectComp(rev, projects);
					
			//删除中间文件
			File file = new File(fullFileName);
			if(file.isFile()) {
				file.delete();
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// 用递归算法获取所有的机器人点焊工序B8_BIWDiscreteOP
	public void getBIWDiscreteOP(TCComponentBOMLine parent) {
		try {
			AIFComponentContext[] children = parent.getChildren();
			//String parentName = parent.getProperty("bl_rev_object_name");
			for (AIFComponentContext chld : children) {
				TCComponentItemRevision rev = ((TCComponentBOMLine) chld.getComponent()).getItemRevision();
				// 如果是点焊工序
				if (rev.isTypeOf("B8_BIWDiscreteOPRevision")) {
					String objectname = Util.getProperty(rev, "object_name");
					if(objectname.length()>0 && "R".equals(objectname.subSequence(0, 1))) {
						list.add((TCComponentBOMLine) chld.getComponent());
						continue;
					}				
				} else {
					getBIWDiscreteOP((TCComponentBOMLine) chld.getComponent());
				}
			}
		} catch (TCException e) {
			// TODO 自动生成的 catch 块
			e.printStackTrace();
		}
	}

	SimpleDateFormat df = new SimpleDateFormat("yyyy/MM/dd");// 设置日期格式

	// 循环点焊工序获取机器人、焊点等信息
	public void getStationBomLine(ArrayList<TCComponentBOMLine> parent) {
		try {
			for (int i = 0; i < parent.size(); i++) {
				AIFComponentContext[] children = parent.get(i).getChildren();
				String parentName = parent.get(i).getProperty("bl_rev_object_name");
				Discrete = new String[11];
				Discrete[0] = parentName;// 点焊工序名称
				Discrete[2] = parentName;// 打顺表中“机器人编号”一栏信息与sheet页名称相同  20191008
				Discrete[1] = df.format(new Date());// 日期
				Discrete[6] = vehicle;// 车型
				int num = 0;// 序号
				if (!beginname.equals(parentName)) {
					num = 0;
					beginname = parentName;
				}

				// 获取焊装工位，先通过点焊工序获取上一级，根据上一级获取焊装工位
				Discrete[5] = getStation(parent.get(i).parent());

				for (AIFComponentContext chld : children) {
					TCComponentItemRevision rev = ((TCComponentBOMLine) chld.getComponent()).getItemRevision();
					System.out.println("对象类型：" + rev.getType());
					// 机器人
					if (rev.isTypeOf("B8_BIWRobotRevision")) {
//						Discrete[2] = Util.getProperty(((TCComponentBOMLine) chld.getComponent()),
//								"bl_rev_object_name");// 机器人编号
						Discrete[3] = Util.getProperty(rev, "b8_Model");// 机器人型号
					}
					// 焊枪
					if (rev.isTypeOf("B8_BIWGunRevision")) {
						Discrete[4] = Util.getProperty(rev, "b8_Model");// 焊枪枪号
					}
					// 焊点
					if (rev.isTypeOf("WeldPointRevision")) {
						Weld = new String[8];
						Weld[0] = parentName;// 点焊工序名称
						Weld[1] = Integer.toString(num + 1);// 序号
						Weld[2] = Util.getProperty(rev, "object_name");// 焊点ID
						Weld[5] = parent.get(i).parent().getProperty("bl_rev_object_name"); //工位名称
						Weld[6] =  parent.get(i).getProperty("bl_sequence_no"); //查询编号
						Weld[7] =  parent.get(i).getProperty("bl_rev_item_id"); //点焊工序ID
						Weldlist.add(Weld);
						num++;
					}

				}
				Discrete[7] = Integer.toString(num);// 焊点的数量
				Discrete[8] = parent.get(i).parent().getProperty("bl_rev_object_name"); //工位名称
				Discrete[9] = parent.get(i).getProperty("bl_sequence_no"); //查询编号
				Discrete[10] = parent.get(i).getProperty("bl_rev_item_id"); //点焊工序ID
				Discretelist.add(Discrete);
			}
		} catch (TCException e) {
			// TODO 自动生成的 catch 块
			e.printStackTrace();
		}
	}

	// 获取焊装工位
	public String getStation(TCComponentBOMLine parent) {
		String stationname = "";
		try {

			AIFComponentContext[] children = parent.getChildren();
			String parentName = parent.getProperty("bl_rev_object_name");

			for (AIFComponentContext chld : children) {
				TCComponentItemRevision rev = ((TCComponentBOMLine) chld.getComponent()).getItemRevision();
				if (rev.isTypeOf("B8_StationRevision")) {
					stationname = Util.getProperty(((TCComponentBOMLine) chld.getComponent()), "bl_rev_object_name");// 焊装工位
					continue;
				}
			}
			return stationname;
		} catch (TCException e) {
			// TODO 自动生成的 catch 块
			e.printStackTrace();
		}
		return stationname;
	}
}
