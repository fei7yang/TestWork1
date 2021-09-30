package com.dfl.report.WeldingParameters;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;

import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

import com.dfl.report.home.OpenHomeDialog;
import com.dfl.report.util.FileUtil;
import com.dfl.report.workschedule.EditionDialog;
import com.dfl.report.workschedule.EngineeringWorkListCoverOp;
import com.teamcenter.rac.aif.AIFDesktop;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class WeldingParametersHandler extends AbstractHandler {
	private AbstractAIFUIApplication app;
	private Shell shell;
	private String Edition;
	private TCComponentFolder rootFolder;
	private TCComponent savefolder;
	private TCSession session;
	private TCComponentBOMLine topbomline;
	private InputStream inputStream;
	private ArrayList rule;

	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {
		// TODO Auto-generated method stub
		app = AIFUtility.getCurrentApplication();
		session = (TCSession) app.getSession();
		rule = getSelectStateRule();
		if (rule == null || rule.size() < 1) {
			MessageBox.post("错误：首选项未定义DFL9_get_version_information，请联系系统管理员！", "错误", MessageBox.INFORMATION);
			return null;
		}
		// 查询工位导出模板
		inputStream = null;
		inputStream = FileUtil.getTemplateFile("DFL_Template_WeldingParameters");
		if (inputStream == null) {
			MessageBox.post("错误：没有找到PSW参数汇总表模板，请联系系统管理员添加模板(名称为：DFL_Template_WeldingParameters)！", "错误", MessageBox.INFORMATION);
			return null;
		}

		InterfaceAIFComponent[] ifc = app.getTargetComponents();
		if (ifc.length < 1) {
			MessageBox.post("当前未选择操作对象，请先选择！", "温馨提示", MessageBox.INFORMATION);
			return null;
		}
		if (ifc.length > 1) {
			MessageBox.post("请选择单一的焊装工厂工艺对象！", "温馨提示", MessageBox.INFORMATION);
			return null;
		}
		if (ifc[0] instanceof TCComponentBOMLine) {

		} else {
			MessageBox.post("请选择焊装工厂工艺对象！", "提示", MessageBox.INFORMATION);
			return null;
		}
		topbomline = (TCComponentBOMLine) ifc[0];
		try {
			System.out.println(topbomline.getItemRevision().getType());
			if (!topbomline.getItemRevision().isTypeOf("B8_BIWPlantBOPRevision")) {
				MessageBox.post("请选择焊装工厂工艺对象！", "温馨提示", MessageBox.INFORMATION);
				return null;
			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		shell = AIFDesktop.getActiveDesktop().getShell();
		Display.getDefault().asyncExec(new Runnable() {
			@Override
			public void run() {
				openDialog();
			}
		});

		return null;
	}

	protected void openDialog() {
		// TODO Auto-generated method stub
		EditionDialog dialog = new EditionDialog(shell, SWT.SHELL_TRIM, true,rule);
		dialog.open();

		Edition = dialog.Edition;
		if (Edition== null || Edition.isEmpty()) {
			try {
				inputStream.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			return;
		}
		Display.getDefault().asyncExec(new Runnable() {
			@Override
			public void run() {
				openHoneDialog();
			}
		});
	}

	protected void openHoneDialog() {
		// TODO Auto-generated method stub
		try {
			rootFolder = session.getUser().getHomeFolder();
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		OpenHomeDialog dialog = new OpenHomeDialog(shell, rootFolder, session);
		dialog.open();

		savefolder = dialog.folder;
		System.out.println("文件夹：" + dialog.folder);

		if (dialog.flag) {
			try {
				inputStream.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			return;
		}

		if (savefolder == null) {
			try {
				inputStream.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			return;
		}

		Thread thread = new Thread() {
			public void run() {
				try {
					new WeldingParametersOp(topbomline, session, Edition, savefolder,inputStream);
				} catch (TCException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		};
		thread.start();
	}

	// 查询阶段首选项，获取阶段信息
	private ArrayList getSelectStateRule() {
		ArrayList rule = new ArrayList();
		try {
			AbstractAIFUIApplication app = AIFUtility.getCurrentApplication();
			TCSession session = (TCSession) app.getSession();
			TCPreferenceService preferenceService = session.getPreferenceService();

			String str = preferenceService.getPreferenceDescription("DFL9_get_version_information");
			if (str != null) {
				String[] values = preferenceService.getStringValues("DFL9_get_version_information");
				if (values != null) {
					for (int i = 0; i < values.length; i++) {
						String value = values[i];
						rule.add(value);
					}
				}
			}
			return rule;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return rule;
	}
}
