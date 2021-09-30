package com.dfl.report.WeldingEquipmentEst;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import javax.swing.JOptionPane;

import org.apache.log4j.Logger;
import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

import com.dfl.report.ExcelReader.WeldPointBoardInformation;
import com.dfl.report.ExcelReader.baseinfoExcelReader;
import com.dfl.report.WeldingEstablishment.WeldingEstablishmentOp;
import com.dfl.report.dialog.SelectionSiteNameDialog;
import com.dfl.report.home.OpenHomeDialog;
import com.dfl.report.util.Util;
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

public class WeldingEquipmentEstHandler extends AbstractHandler {

	private AbstractAIFUIApplication app;
	private Shell shell;
	private TCSession session;
	private TCComponentFolder rootFolder;
	private TCComponent savefolder;
	private String reportname;
	private InterfaceAIFComponent[] aifComponents;
	private List<WeldPointBoardInformation> baseinfolist;
	private static Logger logger = Logger.getLogger(WeldingEquipmentEstHandler.class);

	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {
		// TODO Auto-generated method stub
		try {
			app = AIFUtility.getCurrentApplication();
			aifComponents = app.getTargetComponents();
			session = (TCSession) app.getSession();
			if (aifComponents == null || aifComponents.length < 1) {
				MessageBox.post("����ѡ�����", "����", MessageBox.INFORMATION);
				return null;
			}
			for (int i = 0; i < aifComponents.length; i++) {
				if (aifComponents[i] instanceof TCComponentBOMLine) {

				} else {
					MessageBox.post("��ѡ������д��ڲ���BOMLine����", "��ʾ", MessageBox.INFORMATION);
					return null;
				}
				// �ж���ѡ���������
				TCComponentBOMLine topbomline = (TCComponentBOMLine) aifComponents[i];

				// System.out.println(topbomline.getItemRevision().getType());
				if (!topbomline.getItemRevision().isTypeOf("B8_BIWMEProcLineRevision")
						&& !topbomline.getItemRevision().isTypeOf("B8_BIWPlantBOPRevision")) {
					MessageBox.post("��ѡ������д��ڲ��Ǻ�װ�������հ汾�����װ���߹��հ汾����", "��ʾ", MessageBox.INFORMATION);
					return null;
				}
			}

		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		// ��ȡ��ѡ����Note����
		TCPreferenceService ts = session.getPreferenceService();
		if (!ts.isDefinitionExistForPreference("B8_WeldFeasibilityReport")) {
			MessageBox.post("������ѡ��B8_WeldFeasibilityReportδ����,����ϵϵͳ����Ա��", "��ʾ", MessageBox.INFORMATION);
			logger.error("������ѡ��B8_WeldFeasibilityReportδ����");
			return null;
		}
		InputStream inputStream = Util.getReportTempByprefercen(session, "B8_WeldFeasibilityReport", 2);
		if (inputStream == null) {
			MessageBox.post("�����豸������һԪ��ģ�岻���ڣ�����ϵϵͳ����Ա��", "��ʾ", MessageBox.INFORMATION);
			logger.error("�����豸�����Ա���ģ�岻���ڣ�");
			return null;
		}else {
			try {
				inputStream.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		String baseName = "222.������Ϣ";
		TCComponentBOMLine topbomline = (TCComponentBOMLine) aifComponents[0];
		try {
			baseinfolist = getBaseinfomation(topbomline.window().getTopBOMLine(), baseName);
			if(baseinfolist == null || baseinfolist.size()<1) {
				System.out.println("�������ɹ�����ҵ��-������Ϣ��");
				MessageBox.post("�������ɹ�����ҵ��-������Ϣ��", "��ʾ��Ϣ", MessageBox.ERROR);
				return null;
			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		Thread thread = new Thread() {
			public void run() {
				boolean IsContinu = Util.isContinue("��ȷ����ִ���꺸�Ӳ������㣬�Ƿ��������");
				
//				boolean IsContinu1 = true;
//				boolean IsContinu2 = true;
//				
//				TCPreferenceService ts = session.getPreferenceService();
//				
//				// ��ȡ��ѡ����B8_WeldFeasibilityReport_Allowed_Pres_Diff
//				if(IsContinu)
//				{
//					if (!ts.isDefinitionExistForPreference("B8_WeldFeasibilityReport_Allowed_Pres_Diff")) {
//						IsContinu1 = Util.isContinue("��ѡ�B8_WeldFeasibilityReport_Allowed_Pres_Diff��δ���ã��Ƿ���Ĭ��ֵ2000ִ�У�");
//					}
//					else
//					{
//						IsContinu1 = Util.isContinue("�Ƿ�����ѡ�B8_WeldFeasibilityReport_Allowed_Pres_Diff�����õ�ִֵ�У�");
//					}
//					if(IsContinu1)
//					{
//						if (!ts.isDefinitionExistForPreference("B8_WeldFeasibilityReport_Allowed_Pres_Gap")) {
//							IsContinu2 = Util.isContinue("��ѡ�B8_WeldFeasibilityReport_Allowed_Pres_Gap��δ���ã��Ƿ���Ĭ��ֵ1000ִ�У�");
//						}
//						else
//						{
//							IsContinu2 = Util.isContinue("�Ƿ�����ѡ�B8_WeldFeasibilityReport_Allowed_Pres_Gap�����õ�ִֵ�У�");
//						}
//					}
//				}
//											
				if (IsContinu) {
					execute();
				}
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
		System.out.println("�ļ��У�" + dialog.folder);

		if (dialog.flag) {
			return;
		}

		if (savefolder == null) {
			return;
		}

		SelectionSiteNameDialog dialog2 = new SelectionSiteNameDialog(shell, SWT.SHELL_TRIM);
		dialog2.open();

		reportname = dialog2.name;

		if (reportname == null || reportname.isEmpty()) {
			return;
		}

		Thread thread = new Thread() {
			public void run() {
				try {
					new WeldingEquipmentEstOp(session, aifComponents, reportname, savefolder,baseinfolist);
				} catch (TCException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		};
		thread.start();
	}


	/*
	 * ��ȡ������Ϣ����Ϣ
	 */
	private List<WeldPointBoardInformation> getBaseinfomation(TCComponentBOMLine topbl, String procName) {
		List<WeldPointBoardInformation> baseinfolist = new ArrayList<WeldPointBoardInformation>();
		InputStream filein = null;
		try {
			filein = baseinfoExcelReader.getFileinbyreadExcel2(topbl.getItemRevision(), "IMAN_reference", procName);
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		baseinfolist = baseinfoExcelReader.readHDExcel(filein, "xlsx");

		return baseinfolist;
	}
	
}
