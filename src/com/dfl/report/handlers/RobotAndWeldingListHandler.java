package com.dfl.report.handlers;

import java.io.InputStream;

import javax.swing.SwingUtilities;

import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;

import com.dfl.report.util.FileUtil;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class RobotAndWeldingListHandler extends AbstractHandler {

	public RobotAndWeldingListHandler() {
		// TODO Auto-generated constructor stub
	}

	private AbstractAIFUIApplication app;
	private InputStream inputStream;

	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {
		// TODO Auto-generated method stub
		app = AIFUtility.getCurrentApplication();
		TCSession session = (TCSession) this.app.getSession();
		InterfaceAIFComponent[] ifc = app.getTargetComponents();
		if (ifc.length < 1) {
			MessageBox.post("��ǰδѡ�������������ѡ��", "��ܰ��ʾ", MessageBox.INFORMATION);
			return null;
		}
		if (ifc.length > 1) {
			MessageBox.post("��ѡ��һ�ĺ�װ���߶���", "��ܰ��ʾ", MessageBox.INFORMATION);
			return null;
		}
		if (ifc[0] instanceof TCComponentBOMLine) {

		} else {
			MessageBox.post("��ѡ��BOE�еĺ�װ���߶���", "��ʾ", MessageBox.INFORMATION);
			return null;
		}
		TCComponentBOMLine topbomline = (TCComponentBOMLine) ifc[0];
		try {
			System.out.println(topbomline.getItemRevision().getType());
			if (!topbomline.getItemRevision().isTypeOf("B8_LineRevision")) {
				MessageBox.post("��ѡ��BOE�еĺ�װ���߶���", "��ܰ��ʾ", MessageBox.INFORMATION);
				return null;
			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		// �ж��û�����ѡ�����Ƿ���дȨ��
		boolean flag;
		try {
			flag = Util.hasWritePrivilege(session, topbomline.getItemRevision());
			if (!flag) {
				MessageBox.post("�Ե�ǰ��װ����û��дȨ�ޣ�", "��ܰ��ʾ", MessageBox.INFORMATION);
				return null;
			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}		
		// ��ѯ����ģ��                   
		inputStream = FileUtil.getTemplateFile("DFL_Template_RobotAndWeldingList");

		if (inputStream == null) {
			MessageBox.post("����û���ҵ�������&��ǹ�嵥ģ�壬����ϵϵͳ����Ա���ģ��(����Ϊ��DFL_Template_RobotAndWeldingList)", "��ܰ��ʾ", MessageBox.INFORMATION);			
			return null;
		}

		RobotAndWeldingExportAction action = new RobotAndWeldingExportAction(app, null, "",inputStream);
		Thread th = new Thread(action);
		th.start();

//		SwingUtilities.invokeLater(new RobotAndWeldingExportAction(app,null,""));

		return null;
	}
}
