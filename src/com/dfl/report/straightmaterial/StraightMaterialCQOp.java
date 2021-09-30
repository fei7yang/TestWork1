package com.dfl.report.straightmaterial;

import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dfl.report.util.FileUtil;
import com.dfl.report.util.NewOutputDataToExcel;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentGroup;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentUser;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCSession;

public class StraightMaterialCQOp {

	private AbstractAIFUIApplication app;
	SimpleDateFormat df = new SimpleDateFormat("yyyy年MM月");// 设置日期格式
	SimpleDateFormat df2 = new SimpleDateFormat("yyyy.MM");// 设置日期格式
	private ArrayList list = new ArrayList();//数据集合
	private TCSession session ;
	public StraightMaterialCQOp(AbstractAIFUIApplication app) {
		// TODO Auto-generated constructor stub
		this.app = app;
		session = (TCSession) app.getSession();
		initUI();
	}

	private void initUI() {
		// TODO Auto-generated method stub
		try {
			// 获取选中对象
			InterfaceAIFComponent ifc = app.getTargetComponent();

			TCComponentBOMLine topbomline = (TCComponentBOMLine) ifc;

			// 界面显示进度并输出执行步骤
			ReportViwePanel viewPanel = new ReportViwePanel("生成报表");
			viewPanel.setVisible(true);

			viewPanel.addInfomation("开始输出报表...\n", 5, 100);

			// 根据顶层BOP查询所有的焊装产线
			String[] propertys = new String[] { "bl_item_object_type", "bl_item_object_type" };
			String[] values = new String[] { "直材", "BIW Direct Material" };

			// 直材集合
			ArrayList partList = Util.searchBOMLine(topbomline, "OR", propertys, "==", values);
			
			if(partList==null || partList.size()<1) {
				viewPanel.addInfomation("提示：当前产线下没有直材数据\n", 100, 100);
				return;
			}
			viewPanel.addInfomation("正在输出报表...\n", 20, 100);
			// 获取当前编制人及所在部门
			session = (TCSession) app.getSession();
			TCComponentUser user = session.getUser();
			TCComponentGroup group = user.getLoginGroup();
			String Manufacturing = group.getGroupName();
			String[] cover = new String[4];
			cover[0] = "      编制部门：" + Manufacturing;
			cover[1] = "      编制日期：" + df.format(new Date());
			cover[2] = user.getUserName();
			cover[3] = df2.format(new Date());

			// 查询封面导出模板
			InputStream inputStream = FileUtil.getTemplateFile("DFL_Template_StraightMaterialCQ");

			if (inputStream == null) {
				viewPanel.addInfomation("错误：没有找到直材消耗定额清单模板，请先添加模板(名称为：DFL_Template_StraightMaterialCQ)\n", 100, 100);
				return;
			}
			viewPanel.addInfomation("正在输出报表...\n", 40, 100);
			
			XSSFWorkbook book = NewOutputDataToExcel.creatXSSFWorkbook(inputStream);

			NewOutputDataToExcel.writeDataToSheetByGeneral(book, cover);
			
			viewPanel.addInfomation("正在输出报表...\n", 60, 100);
			
			//根据直材数量分页，每18行数据分一个sheet页
			int sheetnum = (partList.size())/18 + 1;
			
			for(int i=0;i<partList.size();i++) {
				TCComponentBOMLine bl = (TCComponentBOMLine) partList.get(i);
				TCComponentItemRevision rev = bl.getItemRevision();
				String str[] = new String[15];
				str[0] = Integer.toString(i+1);//序号
				str[1] = Util.getProperty(bl, "bl_rev_object_name");//材料名称
				str[2] = "";//DFL编号
				String value[] = getNESAndFactoryByDFL(str[2]);
				if(value!=null && value.length>0) {
					str[3] = value[0];//NES编号
					str[13] = value[1];//厂家
				}
				else {
					str[3] = "";//NES编号
					str[13] = "";//厂家
				}
				str[4] = "";//材料用途
				str[5] = "g";//单位
				str[6] = null;//车型1的单台用量
				str[7] = null;//车型2的单台用量
				str[8] = null;//车型3的单台用量
				str[9] = null;//车型4的单台用量
				str[10] = null;//车型5的单台用量
				str[11] = null;//车型6的单台用量
				str[12] = "";//包装
				
				str[14] = "";//使用班组
							
				list.add(str);
			}
			
			NewOutputDataToExcel.creatXSSFWorkbookByData(book,sheetnum);
			
			viewPanel.addInfomation("正在输出报表...\n", 80, 100);
			
			NewOutputDataToExcel.writeStraightDataToSheet(book,sheetnum,cover,list);
			
			NewOutputDataToExcel.exportFile(book, "直材消耗定额清单");
									
			NewOutputDataToExcel.openFile(FileUtil.getReportFileName("直材消耗定额清单"));
			
			viewPanel.addInfomation("正在输出报表...\n", 100, 100);

		} catch (Exception e) {
			e.printStackTrace();
		}
	}
    //根据DFL编号通过直材对应表获取NES编号和厂家
	private String[] getNESAndFactoryByDFL(String code) {
		try {
			String value[]=new String[2];
			TCPreferenceService preferenceService = session.getPreferenceService();

			String str = preferenceService.getPreferenceDescription("B8_QueryStraightMaterialCorrespondTable");
			if (str == null) {
				return value;
			}
			String[] types = preferenceService.getStringValues("B8_QueryStraightMaterialCorrespondTable");
			
			for(int i=0;i<types.length;i++) {
				String temple[] = types[i].split("=");
				if(temple[0].equals(code)) {
					value[0] = temple[1];
					value[1] = temple[2];
					break;
				}
			}
			return value;
		}
		catch(Exception ex) {
			ex.printStackTrace();
		}
		return null;
	}
}
