package com.dfl.report.handlers;

import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import com.dfl.report.util.FileUtil;
import com.dfl.report.util.OutputDataToExcel2;
import com.dfl.report.util.OutputDataToExcel3;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;

public class NewCarryOverListOp {
	private AbstractAIFUIApplication app;
	private TCSession session;
	private TCComponentBOMLine parent;
	ArrayList childList = new ArrayList();
	ArrayList childList1 = new ArrayList();
	private String[] value;// 属性数组
	private ArrayList valuelist = new ArrayList();
	int number = 0;// 数据行数
	String[][] data;
	private TCComponent folder;
	private Map parentMap;
	private InterfaceAIFComponent[] aifComponents;

	SimpleDateFormat dateformat = new SimpleDateFormat("yyyyMMdd  HH");// 设置日期格式
	
	public NewCarryOverListOp(AbstractAIFUIApplication app, TCComponent folder, InterfaceAIFComponent[] aifComponents, TCSession session) {
		// TODO Auto-generated constructor stub
		this.app = app;
		this.folder=folder;
		this.session = session;
		this.aifComponents =aifComponents;
		parentMap = new HashMap();
		initUI();
	}

	private void initUI() {
		// TODO Auto-generated method stub
		try {
			// 界面显示进度并输出执行步骤
			ReportViwePanel viewPanel = new ReportViwePanel("生成报表");
			viewPanel.setVisible(true);

			// 获取选择的对象
			//InterfaceAIFComponent[] aifComponents = app.getTargetComponents();
			if(aifComponents==null) {
				return;
			}
			TCComponentBOMLine topbomline = (TCComponentBOMLine) aifComponents[0];
			TCComponentBOMLine topbl=topbomline.getCachedWindow().getTopBOMLine();

			viewPanel.addInfomation("正在获取模板...\n", 20, 100);
			// 查询并导出模板
			InputStream inputStream = FileUtil.getTemplateFile("DFL_Template_NewCarryOverList");
			System.out.println("inputStream=" + inputStream);

			if (inputStream == null) {
				viewPanel.addInfomation("错误：没有找到新规留用部品清单模板，请先添加模板(名称为：DFL_Template_NewCarryOverList)\n", 100, 100);
				return;
			}
			
			//修改判断是否移动单元的属性 2019-08-20
			String[] propertys = new String[] { "bl_usage_address"};// 是否移动单元
			String[] values = new String[] { "MU"};
			ArrayList partlist = new ArrayList();
			for(InterfaceAIFComponent aif: aifComponents) {
				TCComponentBOMLine bl = (TCComponentBOMLine) aif;
				ArrayList chlist =  Util.searchBOMLine(bl, "OR", propertys, "=", values);
				for(int i=0;i<chlist.size();i++) {
					if(!partlist.contains(chlist.get(i))) {
						partlist.add(chlist.get(i));
					}
				}
			}
//			ArrayList partlist =  Util.searchBOMLine(topbomline, "OR", propertys, "=", values);
//			String[] propertys = new String[] { "B8_NoteIsBiwTrUnit","B8_NoteIsBiwTrUnit" };// 是否移动单元
//			String[] values = new String[] { "是","1" };
//			ArrayList partlist =  Util.searchBOMLine(topbomline, "OR", propertys, "==", values);
			String tempvehicle= topbl.getItemRevision().getProperty("project_ids");// 车型
			String vehicle = Util.getDFLProjectIdVehicle(tempvehicle);
			if(vehicle==null || vehicle.isEmpty()) {
				vehicle = tempvehicle;
			}
//			System.out.println(vehicle);
			getBomLine(partlist);//获取所有父级对象



			int index = 1;
			Iterator iterator = parentMap.entrySet().iterator();
			while (iterator.hasNext()) 
			{
				Map.Entry entry = (Entry) iterator.next();
				TCComponentBOMLine parent = (TCComponentBOMLine) entry.getKey();
				List childList = (List) entry.getValue();

				
				String parentName = parent.getProperty("bl_rev_object_name");// 零件父级对象的名称
				
				if(childList.isEmpty())
				{
					continue;
				}
				
				String NO = Integer.toString(index);// 零件父级对象的序号；
				
				index++;
				for (int j = 0; j < childList.size(); j++) {
					//判断移动单元是否存在打包
					TCComponentBOMLine packLines[] = ((TCComponentBOMLine) childList.get(j)).getPackedLines();
					if(packLines!=null&&packLines.length>0){
						String counts=((TCComponentBOMLine)childList.get(j)).getProperty("bl_pack_count");
						int count= Integer.parseInt(counts);
						for(int p=count-1;p<count;p++) {
							//String Name =((TCComponentBOMLine)childList.get(j)).getProperty("bl_rev_object_name");// 零件名称
							String Name =Util.getProperty(((TCComponentBOMLine)childList.get(j)).getItemRevision(), "dfl9_CADObjectName");// 零件名称
							String ID = ((TCComponentBOMLine) childList.get(j)).getProperty("bl_DFL9SolItmPartRevision_dfl9_part_no");// 零件号
							String NewCaOver = ((TCComponentBOMLine) childList.get(j)).getProperty("DFL9_new_caover_mark");// 新设/留用

							value = new String[5];

							value[0] = NO;
							value[1] = ID;
							value[2] = Name;
							value[3] = NewCaOver;
							value[4] = parentName;

							valuelist.add(value);

							number++;
						}
						
					}else {
						//String Name =((TCComponentBOMLine)childList.get(j)).getProperty("bl_rev_object_name");// 零件名称
						String Name =Util.getProperty(((TCComponentBOMLine)childList.get(j)).getItemRevision(), "dfl9_CADObjectName");// 零件名称
						String ID = ((TCComponentBOMLine) childList.get(j)).getProperty("bl_DFL9SolItmPartRevision_dfl9_part_no");// 零件号
						String NewCaOver = ((TCComponentBOMLine) childList.get(j)).getProperty("DFL9_new_caover_mark");// 新设/留用

						value = new String[5];

						value[0] = NO;
						value[1] = ID;
						value[2] = Name;
						value[3] = NewCaOver;
						value[4] = parentName;

						valuelist.add(value);

						number++;
					}
					
				}

			}

			viewPanel.addInfomation("开始输出报表...\n", 40, 100);

			// 定义报表行数据
			data = new String[number][13];
			// System.out.println("data=" + data.length);
			for (int i = 0; i < valuelist.size(); i++) {

				String[] str = (String[]) valuelist.get(i);

				data[i][0] = "";// 模板表头前空一格
				data[i][1] = "";// 车型
				data[i][2] = str[0];// 第一个NO
				data[i][3] = str[4];// 块号BLOCK 零件父级对象名称
				data[i][4] = Integer.toString(i + 1);// 第二个NO
				data[i][5] = str[1];// 零件号 Parts No
				data[i][6] = str[2];// 零件名称 Parts Name
				data[i][7] = str[3];// 新设/留用 New/Carry Over
				data[i][8] = "";// ACL
				data[i][9] = str[3];// Datum 跟New/Carry Over一致 B8_NoteNewCaOverMark
				data[i][10] = "";// PosiMark
				data[i][11] = "";// ID Mark
				data[i][12] = "";// Remark
			}
			viewPanel.addInfomation("开始写数据，请耐心等待...\n", 70, 100);
			//输出文件名称
			//String name=Util.getProperty(topbomline.getItemRevision(), "object_name");
			// 文件名称
			String functionname = "";
			for (InterfaceAIFComponent aif : aifComponents) {
				TCComponentBOMLine aifbl = (TCComponentBOMLine) aif;
				if (functionname.isEmpty()) {
					functionname = Util.getProperty(aifbl, "bl_rev_object_name");
				} else {
					functionname = functionname + "&" + Util.getProperty(aifbl, "bl_rev_object_name");
				}
			}
			String date = dateformat.format(new Date());
			String datasetName = vehicle+"_"+functionname+"_"+"新设及留用部品清单"+"_"+date+"时";
			String fileName = Util.formatString(datasetName);
			
			OutputDataToExcel2 outdata = new OutputDataToExcel2(data, inputStream, fileName.trim(), vehicle);
			//存储报表
			Util.saveFiles(fileName.trim(),datasetName, folder, session,"AL");
			
			viewPanel.addInfomation("输出报表完成，请在选择保存的文件夹下查看！\n", 100, 100);
		} catch (Exception e) {
			e.printStackTrace();
		}

	}


	// 获取所有是移动单元的bomline的父级对象
	public void getBomLine(ArrayList partlist) throws TCException {


		for (int i = 0; i < partlist.size(); i++) {
			try {
				TCComponentBOMLine childLine =	(TCComponentBOMLine) partlist.get(i);
				// 获取零件的父级对象
				parent = ((TCComponentBOMLine) partlist.get(i)).parent();

				if(parentMap.containsKey(parent))
				{
					List list = (List) parentMap.get(parent);
					list.add(childLine);
					parentMap.put(parent, list);
				}else
				{
					List list = new ArrayList();
					list.add(childLine);
					parentMap.put(parent, list);
				}

			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}


		System.out.println("parent size="+parentMap.size());

	}

}
