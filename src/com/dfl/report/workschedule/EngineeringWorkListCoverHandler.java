package com.dfl.report.workschedule;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

import com.dfl.report.ExcelReader.CoverInfomation;
import com.dfl.report.ExcelReader.baseinfoExcelReader;
import com.dfl.report.handlers.StraightforwardlistAction;
import com.dfl.report.home.OpenHomeDialog;
import com.dfl.report.util.FileUtil;
import com.dfl.report.util.GenerateReportInfo;
import com.dfl.report.util.ReportUtils;
import com.teamcenter.rac.aif.AIFDesktop;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class EngineeringWorkListCoverHandler extends AbstractHandler {

	private AbstractAIFUIApplication app;
	private Shell shell;
	private String Edition;
	private String LastEdition;
	private GenerateReportInfo info;
	private InputStream inputStream = null;
	private ArrayList rule;

	public EngineeringWorkListCoverHandler() {
		// TODO Auto-generated constructor stub
	}

	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {
		// TODO Auto-generated method stub
		app = AIFUtility.getCurrentApplication();
		InterfaceAIFComponent[] ifc = app.getTargetComponents();
		if (ifc.length < 1) {
			MessageBox.post("当前未选择操作对象，请先选择！", "温馨提示", MessageBox.INFORMATION);
			return null;
		}
		if (ifc.length > 1) {
			MessageBox.post("请选择单一焊装工厂工艺对象！", "温馨提示", MessageBox.INFORMATION);
			return null;
		}
		if (ifc[0] instanceof TCComponentBOMLine) {

		} else {
			MessageBox.post("请选择焊装工厂工艺对象！", "提示", MessageBox.INFORMATION);
			return null;
		}
		TCComponentBOMLine topbomline = (TCComponentBOMLine) ifc[0];
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

		// 文件名称
		String procName = "00.封面";

		// 生成报表操作前的动作
		info = new GenerateReportInfo();
		info.setExist(false);
		info.setIsgoon(true);
		info.setAction(""); //$NON-NLS-1$
		info.setMeDocument(null);
		info.setDFL9_process_type("H"); //$NON-NLS-1$
		info.setDFL9_process_file_type("AB"); // $NON-NLS-1$
		info.setmeDocumentName(procName);

		try {
			info = ReportUtils.beforeGenerateReportAction(topbomline.getItemRevision(), info);
		} catch (TCException e) {
			e.printStackTrace();
			// EclipseUtils.info("Error : " + e.getMessage()); //$NON-NLS-1$
			return null;
		}
		System.out.println("The action is completed before the report operation is generated.");

		if (!info.isIsgoon()) {
			return null;
		}
		if (info.getAction() == "create") {
			inputStream = FileUtil.getTemplateFile("DFL_Template_EngineeringWorkListCover");

			if (inputStream == null) {
				MessageBox.post("错误：没有找到工程作业表封面模板，请联系系统管理员添加模板(名称为：DFL_Template_EngineeringWorkListCover)！", "提示",
						MessageBox.INFORMATION);
				return null;
			}
		} else {
			TCComponentItemRevision docmentRev = info.getMeDocument();
			inputStream = baseinfoExcelReader.getFileinbyreadExcel(docmentRev, "IMAN_specification", procName);

			if (inputStream == null) {
				MessageBox.post("请确认00.封面文档版本对象下，存在00.封面数据集！", "温馨提示", MessageBox.INFORMATION);
				return null;
			}
		}
		rule = getSizeRule();
		if(rule == null || rule.size()<1) {
			MessageBox.post("错误：首选项DFL9_get_version_information未定义，请联系系统管理员！", "温馨提示", MessageBox.INFORMATION);
			return null;
		}
		
		List<CoverInfomation> list = getCoverinfomation(topbomline, "00.封面");
		if (list != null && list.size() > 0) {
			CoverInfomation cif = list.get(0);
			LastEdition = cif.getEdition();
		} else {
			LastEdition = "";
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
		EditionDialog dialog = new EditionDialog(shell, SWT.SHELL_TRIM, LastEdition,rule);
		dialog.open();

		Edition = dialog.Edition;
		if (Edition == null || Edition.isEmpty()) {
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
					new EngineeringWorkListCoverOp(app, null, Edition,info,inputStream);
				} catch (TCException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		};
		thread.start();

	}

	/*
	 * 获取封面信息信息
	 */
	private List<CoverInfomation> getCoverinfomation(TCComponentBOMLine topbl, String procName) {
		List<CoverInfomation> coverinfolist = new ArrayList<CoverInfomation>();
		InputStream filein = null;
		try {
			filein = baseinfoExcelReader.getFileinbyreadExcel2(topbl.getItemRevision(), "IMAN_reference", procName);
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		coverinfolist = baseinfoExcelReader.readCoverExcel(filein, "xlsx");

		return coverinfolist;
	}

	// 查询版次首选项，获取版次信息
	private ArrayList getSizeRule() {
		ArrayList rule = new ArrayList();
		try {
			AbstractAIFUIApplication app = AIFUtility.getCurrentApplication();
			TCSession session = (TCSession) app.getSession();
			TCPreferenceService preferenceService = session.getPreferenceService();

			String str = preferenceService.getPreferenceDescription("DFL9_get_version_information");
			if (str != null) {
				String[] values = preferenceService.getStringValues("DFL9_get_version_information");
				if(values !=null) {
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
