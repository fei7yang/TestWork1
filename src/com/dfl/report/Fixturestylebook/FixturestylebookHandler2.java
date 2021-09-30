package com.dfl.report.Fixturestylebook;

import java.io.IOException;
import java.io.InputStream;
import java.rmi.AccessException;
import java.util.ArrayList;
import java.util.List;

import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

import com.dfl.report.DotStatistics.DotStatisticsOp;
import com.dfl.report.util.FileUtil;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AIFDesktop;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class FixturestylebookHandler2 extends AbstractHandler {

	private AbstractAIFUIApplication app;
	private Shell shell;
	private TCSession session;
	private TCComponent savefolder;
	// private String page1;
	private String page2;
	private InterfaceAIFComponent[] aifComponents;
	private ArrayList list = new ArrayList();
	private String isupdateflag;//1：终止；2：不更新；3更新

	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {
		// TODO Auto-generated method stub
		try {
			app = AIFUtility.getCurrentApplication();
			session = (TCSession) app.getSession();
			aifComponents = app.getTargetComponents();
			if (aifComponents == null || aifComponents.length < 1) {
				MessageBox.post("请先选择对象！", "错误", MessageBox.INFORMATION);
				return null;
			}
			if (aifComponents.length > 1) {
				MessageBox.post("请选择单一的焊装产线工艺对象或焊装工位工艺对象！", "错误", MessageBox.INFORMATION);
				return null;
			}
			if (aifComponents[0] instanceof TCComponentBOMLine) {

			} else {
				MessageBox.post("所选择对象中存在不是BOMLine对象！", "提示", MessageBox.INFORMATION);
				return null;
			}
			// 判断所选对象的类型
			TCComponentBOMLine topbomline = (TCComponentBOMLine) aifComponents[0];

			// System.out.println(topbomline.getItemRevision().getType());
			if (!topbomline.getItemRevision().isTypeOf("B8_BIWMEProcLineRevision")&&!topbomline.getItemRevision().isTypeOf("B8_BIWMEProcStatRevision")) {
				MessageBox.post("所选择对象中存在不是焊装产线工艺版本或焊装工位工艺版本对象！", "提示", MessageBox.INFORMATION);
				return null;
			}
			//判断是否维护报表模板
			InputStream inputStream = FileUtil.getTemplateFile("DFL_Template_FixtureStyleBook");

			if (inputStream == null) {
				MessageBox.post("错误：没有找到夹具式样书模板，请联系系统管理员添加模板(名称为：DFL_Template_FixtureStyleBook)！", "提示", MessageBox.INFORMATION);
				//viewPanel.addInfomation("错误：没有找到夹具式样书模板，请先添加模板(名称为：DFL_Template_FixtureStyleBook)\n", 100, 100);
				return null;
			}else {
				inputStream.close();
			}
			
			String typename = Util.getObjectDisplayName(session, "B8_BIWMEProcStat");
			String[] propertys = new String[] { "bl_item_object_type", "bl_item_object_type" };
			String[] values = new String[] { typename, typename };
			if(topbomline.getItemRevision().isTypeOf("B8_BIWMEProcLineRevision")) {
				// 循环所有工位，判断是否有已生成过报表
				list = Util.searchBOMLine(topbomline, "OR", propertys, "==", values);
			}else {
				if(list!=null&& list.size()>0) {
					list.clear();
				}
				list.add(topbomline);
			}		
			List finishlist = new ArrayList();
			//标记对焊装工位工艺版本对象关系下的夹具式样书文档版本是否有写权限
			String alldocuments = "";
			//标记对焊装工位工艺对象是否有写权限
			String allgw = "";
			for (int i = 0; i < list.size(); i++) {
				// 根据选择的焊装工位工艺下，是否已经生成过报表，如果生成过则直接取之前的报表作为模板
				TCComponentBOMLine gwbl = (TCComponentBOMLine) list.get(i);
				TCComponentItemRevision blrev = gwbl.getItemRevision();
				String gwname = Util.getProperty(blrev, "object_name");
				// 输出的文件名称
				String datasetname = gwname + "夹具式样书";
				String filename = Util.formatString(datasetname);
				TCComponent[] tccs = blrev.getRelatedComponents("IMAN_reference");
				TCComponentItem tcc = null;
				TCComponentItemRevision oldrev = null;
				System.out.println("关系对象数组：" + tccs);
				for (TCComponent item : tccs) {
					String type = Util.getRelProperty(item, "object_name");
					if (type.equals(datasetname)) {
						tcc = (TCComponentItem) item;
						break;
					}
				}
				System.out.println("关系对象：" + tcc);
				if (tcc == null) {
					finishlist.add(gwname);
				}else {
					oldrev = tcc.getLatestItemRevision();
					//判断是否已发布
					if(oldrev.getDateProperty("date_released") != null) {
						if(!tcc.okToModify()) {
							if(allgw.isEmpty()) {
								allgw = gwname;
							}else {
								allgw = allgw + "," + gwname;
							}
						}
					}else {
						// 判断用户对所选对象是否有写权限
						boolean flag1 = Util.hasWritePrivilege(session, oldrev);
						if (!flag1) {					
							if(alldocuments.isEmpty()) {
								alldocuments = gwname;
							}else {
								alldocuments = alldocuments + "," + gwname;
							}
						}	
					}
					
				}
			}
			String gwlist = "";
			if (finishlist != null && finishlist.size() > 0) {
				for (int i = 0; i < finishlist.size(); i++) {
					String name = (String) finishlist.get(i);
					if (gwlist.isEmpty()) {
						gwlist = name;
					} else {
						gwlist = gwlist + "," + name;
					}
				}
				MessageBox.post("存在工位：" + gwlist + "未生成报表，无法更新！", "提示", MessageBox.INFORMATION);
				return null;
			}
			if(!alldocuments.isEmpty()) {
				MessageBox.post("对当前所选" + alldocuments + "焊装工位工艺版本对象关系下的夹具式样书文档版本对象没有写权限！", "提示", MessageBox.INFORMATION);	
				return null;
			}
			if(!allgw.isEmpty()) {
				MessageBox.post("对当前所选" + allgw + "焊装工位工艺版本对象关系下的夹具式样书文档版本已发布，没有权限执行修订！", "提示", MessageBox.INFORMATION);	
				return null;
			}

		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		page2 = "3"; // 默认3页
		Thread thread = new Thread() {
			public void run() {
				
				shell = AIFDesktop.getActiveDesktop().getShell();
				
				Display.getDefault().asyncExec(new Runnable() {
					@Override
					public void run() {
						execute();
					}
				});
//				try {
//					new FixturestylebookOp(session, aifComponents, page2, true,list);
//				} catch (TCException | AccessException e) {
//					// TODO Auto-generated catch block
//					e.printStackTrace();
//				}
			}
		};
		thread.start();

//		Display.getDefault().asyncExec(new Runnable() {
//			@Override
//			public void run() {
//				execute();
//			}
//		});

		return null;
	}

	protected void execute() {
		// TODO Auto-generated method stub
		
		UpdateConfirmDialog dialog = new UpdateConfirmDialog(shell,SWT.SHELL_TRIM);
		dialog.open();
		isupdateflag = dialog.updateflag;
		if("1".equals(isupdateflag))
		{
			return;
		}
		
		Thread thread = new Thread() {
			public void run() {
				try {
					new FixturestylebookOp(session, aifComponents, page2, true,list,isupdateflag);
				} catch (TCException | AccessException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		};
		thread.start();
		
	}
}
