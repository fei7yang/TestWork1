package com.dfl.report.handlers;

import java.io.IOException;
import java.io.InputStream;

import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

import com.dfl.report.home.OpenHomeDialog;
import com.dfl.report.util.FileUtil;
import com.dfl.report.watertight.WatertightListAction;
import com.teamcenter.rac.aif.AIFDesktop;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class NewCarryOverListHandler extends AbstractHandler {

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
			aifComponents = app.getTargetComponents();
			if (aifComponents == null || aifComponents.length < 1) {
				MessageBox.post("请先选择BBOM总成或者BBOM功能分区对象！", "错误", MessageBox.INFORMATION);
				return null;
			}

			for (InterfaceAIFComponent aif : aifComponents) {
				if (aif instanceof TCComponentBOMLine) {

				} else {
					MessageBox.post("请选择BBOM总成或者BBOM功能分区对象进行该操作！", "提示", MessageBox.INFORMATION);
					return null;
				}
				// 判断所选对象的类型
				TCComponentBOMLine topbomline = (TCComponentBOMLine) aif;
				String type = topbomline.getItemRevision().getType();
				System.out.println(type);
				if ((!type.equals("B8_BBOMTopNodeRevision")) && (!type.equals("B8_BBOMPartitionRevision"))) {
					MessageBox.post("请选择BBOM总成或者BBOM功能分区对象进行该操作！", "提示", MessageBox.INFORMATION);
					return null;
				}
			}

		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		// 查询并导出模板
		InputStream inputStream = FileUtil.getTemplateFile("DFL_Template_NewCarryOverList");
		System.out.println("inputStream=" + inputStream);

		if (inputStream == null) {
			MessageBox.post("错误：没有找到新规留用部品清单模板，请联系系统管理员添加模板(名称为：DFL_Template_NewCarryOverList)！", "提示", MessageBox.INFORMATION);
			return null;
		}else {
			try {
				inputStream.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
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
		session = (TCSession) app.getSession();
		shell = AIFDesktop.getActiveDesktop().getShell();

		// InterfaceAIFComponent aifComponent = app.getTargetComponent();

		try {
			rootFolder = session.getUser().getHomeFolder();
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		// rootFolder = (TCComponent) aifComponent;

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
		NewCarryOverListAction action = new NewCarryOverListAction(app, null, savefolder, aifComponents, session);
		new Thread(action).start();
	}

}
