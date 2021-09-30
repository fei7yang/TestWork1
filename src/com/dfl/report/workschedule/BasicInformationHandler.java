package com.dfl.report.workschedule;

import java.io.IOException;
import java.io.InputStream;

import javax.swing.JOptionPane;

import org.apache.poi.ss.formula.functions.T;
import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.swt.widgets.Display;
import org.eclipse.ui.PlatformUI;

import com.dfl.report.ExcelReader.baseinfoExcelReader;
import com.dfl.report.util.FileUtil;
import com.dfl.report.util.GenerateReportInfo;
import com.dfl.report.util.ReportUtils;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AIFDesktop;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.ConfirmDialog;
import com.teamcenter.rac.util.MessageBox;

public class BasicInformationHandler extends AbstractHandler {

	private AbstractAIFUIApplication app;
    private GenerateReportInfo info ;
    private InputStream inputStream = null;
	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {
		// TODO Auto-generated method stub

		app = AIFUtility.getCurrentApplication();
		InterfaceAIFComponent[] ifc = app.getTargetComponents();
		if (ifc.length < 1) {
			MessageBox.post("��ǰδѡ�������������ѡ��", "��Ϣ��ʾ", MessageBox.INFORMATION);
			return null;
		}
		if (ifc.length > 1) {
			MessageBox.post("��ѡ��һ�ĺ�װ�������ն���", "��Ϣ��ʾ", MessageBox.INFORMATION);
			return null;
		}
		if (ifc[0] instanceof TCComponentBOMLine) {

		} else {
			MessageBox.post("��ѡ��װ�������ն���", "��Ϣ��ʾ", MessageBox.INFORMATION);
			return null;
		}
		TCComponentBOMLine topbomline = (TCComponentBOMLine) ifc[0];
		try {
			System.out.println(topbomline.getItemRevision().getType());
			if (!topbomline.getItemRevision().isTypeOf("B8_BIWPlantBOPRevision")) {
				MessageBox.post("��ѡ��װ�������ն���", "��ܰ��ʾ", MessageBox.INFORMATION);
				return null;
			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		/******************* �жϷŵ�ǰ�� *******************************/
		// �ļ�����
		String procName = "222.������Ϣ";
		// ���ɱ������ǰ�Ķ���
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
		
		if (info.getAction() == "create") { // �����
			// ��ѯ����ģ��
			inputStream = FileUtil.getTemplateFile("DFL_Template_BasicInformation");
			if (inputStream == null) {
				MessageBox.post("����û���ҵ�������Ϣ��ģ�壬�������ģ��(����Ϊ��DFL_Template_BasicInformation)������ϵϵͳ����Ա��", "��Ϣ��ʾ", MessageBox.INFORMATION);
				return null;
			}

		} else {
			TCComponentItemRevision docmentRev = info.getMeDocument();
			inputStream = baseinfoExcelReader.getFileinbyreadExcel(docmentRev, "IMAN_specification", procName);

			if (inputStream == null) {
				MessageBox.post("��ȷ��222.������Ϣ�ĵ��汾�����£�����222.������Ϣ���ݼ���", "��Ϣ��ʾ", MessageBox.INFORMATION);
				return null;
			} 
		}
		/*************************************************/
/*		System.out.println("���ڲ�������Ӧ����");

		Thread thread = new Thread() {
			public void run() {
				boolean IsContinu = Util.isContinue("�Ḳ����һ������ĺ�����Ϣ����ȷ���Ƿ�����������");
				if (IsContinu) {
					try {
						new BasicInformationOp(app,info,inputStream);
					} catch (TCException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}else {
					try {
						inputStream.close();
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
			}
		};
		thread.start();
*/
		
		ConfirmDialog cDlg = new ConfirmDialog( PlatformUI.getWorkbench().getActiveWorkbenchWindow().getShell(), "", "�Ḳ����һ������ĺ�����Ϣ����ȷ���Ƿ�����������",true );
        cDlg.open();

        if( cDlg.isOkayClicked() )
        {
        	try {
				BasicInformationOp op = new BasicInformationOp(app,info,inputStream);
				op.setStartMessage("���������Ϣ...");
				((TCSession) app.getSession()).queueOperation(op);
			} catch (TCException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
        }else {
        	try {
				inputStream.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		
		
		return null;
	}
}
