package com.dfl.report.workschedule;

import java.io.IOException;
import java.io.InputStream;
import java.rmi.AccessException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

import com.dfl.report.ExcelReader.CoverInfomation;
import com.dfl.report.ExcelReader.CurrentandVoltage;
import com.dfl.report.ExcelReader.WeldPointBoardInformation;
import com.dfl.report.ExcelReader.baseinfoExcelReader;
import com.dfl.report.handlers.AntirustRequirementsCheckAction;
import com.dfl.report.home.OpenHomeDialog;
import com.dfl.report.util.FileUtil;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AIFDesktop;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class StationInformationTableHandler extends AbstractHandler {

	private AbstractAIFUIApplication app;
	private Shell shell;
	private LinkedHashMap<String, String> map = new LinkedHashMap<String, String>();
	private String nameNO;
	private String Edition;
	private String topfoldername;
	private String model;
	private boolean IsSameout;
	private ArrayList list = new ArrayList();
	private TCSession session;
	private List<CurrentandVoltage> cv;
	private List<WeldPointBoardInformation> baseinfolist;
	private Map<String,List<String>> MaterialMap;

	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {
		// TODO Auto-generated method stub
		app = AIFUtility.getCurrentApplication();
		session = (TCSession) app.getSession();
		InterfaceAIFComponent[] ifc = app.getTargetComponents();
		if (ifc.length < 1) {
			MessageBox.post("��ǰδѡ�������������ѡ��", "��ʾ��Ϣ", MessageBox.INFORMATION);
			return null;
		}
		if (ifc.length > 1) {
			MessageBox.post("��ѡ��һ��װ��λ���ն���", "��ʾ��Ϣ", MessageBox.INFORMATION);
			return null;
		}
		if (ifc[0] instanceof TCComponentBOMLine) {

		} else {
			MessageBox.post("��ѡ��װ��λ���ն���", "��ʾ", MessageBox.INFORMATION);
			return null;
		}
		TCComponentBOMLine topbomline = (TCComponentBOMLine) ifc[0];

		try {
			System.out.println(topbomline.getItemRevision().getType());
			if (!topbomline.getItemRevision().isTypeOf("B8_BIWMEProcStatRevision")) {
				MessageBox.post("��ѡ��װ��λ���ն���", "��ʾ��Ϣ", MessageBox.INFORMATION);
				return null;
			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		// ��ȡ��ѡ����Note����
		TCPreferenceService ts = session.getPreferenceService();
		if (!ts.isDefinitionExistForPreference("B8_Calculation_Parameter_Name")) {
			MessageBox.post("������ѡ��B8_Calculation_Parameter_Nameδ����,����ϵϵͳ����Ա��", "��ʾ��Ϣ", MessageBox.ERROR);
			return null;
		}
		//��ȡ���϶��ձ�
		MaterialMap = baseinfoExcelReader.getMaterialComparisonTable(app, "DFL_MaterialMapping");
		if(MaterialMap == null || MaterialMap.size()<1)
		{
			System.out.println("δ�ҵ����϶��ձ�");
			MessageBox.post("δ���ö��ձ�DFL_MaterialMapping������ϵϵͳ����Ա��", "��ʾ��Ϣ", MessageBox.ERROR);
			return null;
		}
		// ��ȡ�������
		Object[] obj = baseinfoExcelReader.getCalculationParameter(app, "B8_Calculation_Parameter_Name");
		cv = new ArrayList<CurrentandVoltage>();
		if (obj != null) {
			if (obj[1] != null) {
				cv = (List<CurrentandVoltage>) obj[1];
			} else {
				System.out.println("δ�ҵ����Ӳ����������");
				MessageBox.post("����δ�ҵ����Ӳ����������", "��ʾ��Ϣ", MessageBox.ERROR);
				return null;
			}
		}else {
			System.out.println("δ�ҵ����Ӳ����������");
			MessageBox.post("����δ�ҵ����Ӳ����������", "��ʾ��Ϣ", MessageBox.ERROR);
			return null;
		}
		Map<String, String> map =  getSizeRule();
		if(map == null || map.size()<1) {
			System.out.println("��ѡ��DFL9_get_parts_source δ���ã�����ϵϵͳ����Ա��");
			MessageBox.post("������ѡ��DFL9_get_parts_source δ���ã�����ϵϵͳ����Ա��", "��ʾ��Ϣ", MessageBox.ERROR);
			return null;
		}
		InputStream inputStream = null;
		inputStream = FileUtil.getTemplateFile("DFL_Template_EngineeringWorkListStation");
		if (inputStream == null) {
			MessageBox.post("����û���ҵ�������ҵ����ͨ��λģ�壬����ϵϵͳ����Ա���ģ��(����Ϊ��DFL_Template_EngineeringWorkListStation)��", "��ʾ��Ϣ", MessageBox.ERROR);				
			return null;
		}
		if(inputStream != null) {
			try {
				inputStream.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		inputStream = FileUtil.getTemplateFile("DFL_Template_EngineeringWorkVINCarve");
		if (inputStream == null) {
			MessageBox.post("����û���ҵ�������ҵ��VIN����ģ�壬����ϵϵͳ����Ա���ģ��(����Ϊ��DFL_Template_EngineeringWorkVINCarve)��", "��ʾ��Ϣ", MessageBox.ERROR);				
			return null;
		}
		if(inputStream != null) {
			try {
				inputStream.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		inputStream = FileUtil.getTemplateFile("DFL_Template_AdjustmentLine");
		if (inputStream == null) {
			MessageBox.post("����û���ҵ�������ҵ�������ģ�壬����ϵϵͳ����Ա���ģ��(����Ϊ��DFL_Template_AdjustmentLine)��", "��ʾ��Ϣ", MessageBox.ERROR);				
			return null;
		}	
		if(inputStream != null) {
			try {
				inputStream.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		
		String baseName = "222.������Ϣ";
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
		List<CoverInfomation> list;
		try {
			list = getCoverinfomation(topbomline.window().getTopBOMLine(), "00.����");
			if (list != null && list.size() > 0) {
				CoverInfomation cif = list.get(0);
				Edition = cif.getEdition();
				topfoldername = cif.getFilecode();
			} else {
				System.out.println("�������ɹ�����ҵ��-���棡");
				MessageBox.post("�������ɹ�����ҵ��-���棡", "��ʾ��Ϣ", MessageBox.ERROR);
				return null;
			}
		} catch (TCException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}

		Thread thread = new Thread() {
			public void run() {
				boolean IsContinu = Util.isContinue("��������ɹ������ٴ�����Ḳ��֮ǰ�����ɵı�����ȷ���Ƿ�����������");
				if (!IsContinu) {
					return;
				}
				execute();
			}
		};
		thread.start();

		return null;
	}

	protected void execute() {
		Display.getDefault().asyncExec(new Runnable() {
			@Override
			public void run() {

				OpenDialog();
			}
		});
	}

	protected void OpenDialog() {
		// TODO Auto-generated method stub

		shell = AIFDesktop.getActiveDesktop().getShell();

		SelectTemplate dialog = new SelectTemplate(shell, SWT.SHELL_TRIM);
		dialog.open();

		map = dialog.map;
		list = dialog.list;
		model = dialog.model;
		nameNO = dialog.nameNO;
		IsSameout = dialog.IsSameout;
		System.out.println("sheet������" + map.size());
		if (map == null || map.size() < 1) {
			return;
		}
		Thread thread = new Thread() {
			public void run() {
				try {
					new StationInformationTableOp(app, list, map, Edition, model, nameNO, topfoldername, IsSameout,cv,baseinfolist,MaterialMap);
				} catch (TCException | AccessException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		};
		thread.start();

	}

	/*
	 * ��ȡ������Ϣ��Ϣ
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

	// ��ѯ��Ʒ������ѡ���ȡ��Ʒ������Ϣ
	private Map<String, String> getSizeRule() {
		Map<String, String> rule = new HashMap<String, String>();
		try {
			AbstractAIFUIApplication app = AIFUtility.getCurrentApplication();
			TCSession session = (TCSession) app.getSession();
			TCPreferenceService preferenceService = session.getPreferenceService();

			String str = preferenceService.getPreferenceDescription("DFL9_get_parts_source");
			if (str != null) {
				String[] values = preferenceService.getStringValues("DFL9_get_parts_source");
				for (int i = 0; i < values.length; i++) {
					String value = values[i];
					if (value != null) {
						String[] val = value.split("=");
						if (val != null && val.length > 1) {
							rule.put(val[0], val[1]);
						}
					}
				}
			}
			return rule;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return rule;
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
