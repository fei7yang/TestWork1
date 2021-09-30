package com.dfl.report.DotStatistics;

import java.io.IOException;
import java.io.InputStream;

import org.apache.xalan.xsltc.compiler.util.Util;
import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

import com.dfl.report.home.OpenHomeDialog;
import com.dfl.report.util.FileUtil;
import com.teamcenter.rac.aif.AIFDesktop;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentGroup;
import com.teamcenter.rac.kernel.TCComponentUser;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class DotStatisticsHandler extends AbstractHandler {

	private AbstractAIFUIApplication app;
	private Shell shell;
	private TCSession session;
	private TCComponentFolder rootFolder;
	private TCComponent savefolder;
	private InterfaceAIFComponent[] aifComponents;

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
				MessageBox.post("请选择单一的焊装工厂工艺对象！", "错误", MessageBox.INFORMATION);
				return null;
			}

			if (aifComponents[0] instanceof TCComponentBOMLine) {

			} else {
				MessageBox.post("选择对象不是BOMLine对象！", "提示", MessageBox.INFORMATION);
				return null;
			}
			// 判断所选对象的类型
			TCComponentBOMLine topbomline = (TCComponentBOMLine) aifComponents[0];

			// System.out.println(topbomline.getItemRevision().getType());
			if (!topbomline.getItemRevision().isTypeOf("B8_BIWPlantBOPRevision")) {
				MessageBox.post("请选择焊装工厂工艺对象！", "提示", MessageBox.INFORMATION);
				return null;
			}
			// 查询目录导出模板
			InputStream inputStream = FileUtil.getTemplateFile("DFL_Template_DotStatistics");

			if (inputStream == null) {
				MessageBox.post("错误：没有找到打点统计表模板，请联系系统管理员添加模板(名称为：DFL_Template_DotStatistics)", "提示", MessageBox.INFORMATION);
				return null;
			}else {
				try {
					inputStream.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
//			TCComponentGroup group = session.getGroup();			
//			String groupname = group.getGroupName();
//			
//			if(!groupname.equals("同期工程科")) {
//				MessageBox.post("您不是同期工程科的，无权限生成该报表！", "提示", MessageBox.INFORMATION);
//				return null;
//			}
					
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		Thread thread = new Thread() {
			public void run() {
				execute();
			}
		};
		thread.start();

		return null;
	}

	protected void execute() {
		// TODO Auto-generated method stub
		
		shell = AIFDesktop.getActiveDesktop().getShell();

		try {
			rootFolder = session.getUser().getHomeFolder();
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		Display.getDefault().asyncExec(new Runnable() {
			@Override
			public void run() {
				openDialog();
			}
		});

	}

	protected void openDialog() {
		// TODO Auto-generated method stub
		OpenHomeDialog dialog = new OpenHomeDialog(shell, rootFolder, session);
		dialog.open();

		savefolder = dialog.folder;
		System.out.println("文件夹：" + dialog.folder);

		if (dialog.flag) {
			return;
		}

		if (savefolder == null) {
			return;
		}

		Thread thread = new Thread() {
			public void run() {
				try {
					new DotStatisticsOp(session, aifComponents,savefolder);
				} catch (TCException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		};
		thread.start();
	}
}
