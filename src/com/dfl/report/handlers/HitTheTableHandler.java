package com.dfl.report.handlers;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;

import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;

import com.dfl.report.util.FileUtil;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class HitTheTableHandler extends AbstractHandler {

	public HitTheTableHandler() {
		// TODO Auto-generated constructor stub
	}

	private AbstractAIFUIApplication app;
	private InputStream inputStream;

	@Override
	public Object execute(ExecutionEvent event) throws ExecutionException {
		// TODO Auto-generated method stub
		app = AIFUtility.getCurrentApplication();
		TCSession session = (TCSession) app.getSession();
		InterfaceAIFComponent[] ifc = app.getTargetComponents();
		if (ifc.length < 1) {
			MessageBox.post("当前未选择操作对象，请先选择！", "温馨提示", MessageBox.INFORMATION);
			return null;
		}
		if (ifc.length > 1) {
			MessageBox.post("请选择单一焊装产线工艺对象！", "温馨提示", MessageBox.INFORMATION);
			return null;
		}
		if (ifc[0] instanceof TCComponentBOMLine) {

		} else {
			MessageBox.post("请选择BOP中的焊装产线工艺对象！", "提示", MessageBox.INFORMATION);
			return null;
		}
		TCComponentBOMLine topbomline = (TCComponentBOMLine) ifc[0];
		try {
			System.out.println(topbomline.getItemRevision().getType());
			if (!topbomline.getItemRevision().isTypeOf("B8_BIWMEProcLineRevision")) {
				MessageBox.post("请选择BOP中的焊装产线工艺对象！", "温馨提示", MessageBox.INFORMATION);
				return null;
			}
			TCComponentBOMLine parentline = topbomline.parent();
			if (!parentline.getItemRevision().isTypeOf("B8_BIWMEProcLineRevision")) {
				MessageBox.post("请选择焊装实际产线工艺输出报表！", "温馨提示", MessageBox.INFORMATION);
				return null;
			}
			TCComponentBOMLine topline = topbomline.window().getTopBOMLine();
			TCComponent[] projects = topline.getItemRevision().getRelatedComponents("project_list");
			if(projects == null || projects.length<1)
			{
				MessageBox.post("请将BOP顶层指派项目！", "温馨提示", MessageBox.INFORMATION);
				return null;
			}
			
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		boolean Isupdateflag = true;
		TCComponentItem tcc = null;

		// 根据选择的焊装产线工艺下，是否已经生成过报表，如果生成过则直接取之前的报表作为模板
		TCComponentItemRevision blrev;
		try {
			blrev = topbomline.getItemRevision();
			// 输出的文件名称
			String datasetname = topbomline.getProperty("bl_rev_object_name") + "打顺表";
			//String fileName = Util.formatString(datasetname);
			String fileName = datasetname;
			TCComponent[] tccs = blrev.getRelatedComponents("IMAN_reference");
			
			System.out.println("关系对象数组：" + tccs);
			for (TCComponent item : tccs) {
				String type = Util.getRelProperty(item, "b8_BIWProcDocType");
				System.out.println("工艺文档类类型：" + type);
				if (type.equals("AA") || type.equals("打顺表")) {
					tcc = (TCComponentItem) item;
					break;
				}
			}
			System.out.println("关系对象：" + tcc);

			if (tcc == null) {
				// 判断用户对所选对象是否有写权限
				boolean flag = Util.hasWritePrivilege(session, blrev);
				if (!flag) {
					MessageBox.post("对当前焊装产线工艺没有写权限！", "温馨提示", MessageBox.INFORMATION);
					return null;
				}
				// 查询导出模板
				inputStream = FileUtil.getTemplateFile("DFL_Template_HitTheTable");

				if (inputStream == null) {
					MessageBox.post("错误：没有找到打顺表模板，请联系系统管理员添加模板(名称为：DFL_Template_HitTheTable)", "温馨提示",
							MessageBox.INFORMATION);
					return null;
				}
			} else {
				TCComponentItemRevision rev = tcc.getLatestItemRevision();
				TCComponent[] tccdata = rev.getRelatedComponents("IMAN_specification");
				TCComponentDataset dataset = null;
				File file = null;
				
				// 判断用户对所选对象是否有写权限
				boolean flag = Util.hasWritePrivilege(session, rev);
				if (!flag) {
					MessageBox.post("对当前焊装产线工艺关系下的打顺表文档版本对象没有写权限！", "温馨提示", MessageBox.INFORMATION);
					return null;
				}
				
				if (tccdata != null && tccdata.length > 0) {
					dataset = (TCComponentDataset) tccdata[0];
				}
				System.out.println("获取的数据集：" + dataset);
				if (dataset != null) {
					String filepath = System.getProperty("java.io.tmpdir");
					File tf=new File(filepath+fileName + ".xlsx");
					if(tf.exists())
						tf.delete();
					file = dataset.getFile("excel", fileName + ".xlsx", dataset.getWorkingDir());
				}
				if (file == null) {
					Isupdateflag = false; // 存在文档对象，不存在报表数据集
					System.out.println("测试问题：" + file);
					inputStream = FileUtil.getTemplateFile("DFL_Template_HitTheTable");
					if (inputStream == null) {
						MessageBox.post("错误：没有找到打顺表模板，请联系系统管理员添加模板(名称为：DFL_Template_HitTheTable)", "温馨提示",
								MessageBox.INFORMATION);
						return null;
					}
				} else {
					// 根据获取的报表为模板
					inputStream = new FileInputStream(file);					
					if (inputStream == null) {
						System.out.println("测试问题：" + inputStream);
						Isupdateflag = false; // 存在文档对象，不存在报表数据集
						inputStream = FileUtil.getTemplateFile("DFL_Template_HitTheTable");
						if (inputStream == null) {
							MessageBox.post("错误：没有找到打顺表模板，请联系系统管理员添加模板(名称为：DFL_Template_HitTheTable)", "温馨提示",
									MessageBox.INFORMATION);
							return null;
						}
					}
					else
					{
						System.out.println("获取到已生成的报表：" + Isupdateflag);
					}
				}
				System.out.println("报表文件：" + file);

			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		HitTheTableAction action = new HitTheTableAction(app, null, "",tcc,inputStream,Isupdateflag);
		Thread th = new Thread(action);
		th.start();

		return null;
	}

}
