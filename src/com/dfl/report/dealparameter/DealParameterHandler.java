package com.dfl.report.dealparameter;

import java.awt.Frame;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.swing.JTable;

import org.apache.log4j.Logger;
import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;

import com.dfl.report.ExcelReader.CurrentandVoltage;
import com.dfl.report.ExcelReader.RecommendedPressure;
import com.dfl.report.ExcelReader.SFSequenceWeldingConditionList;
import com.dfl.report.ExcelReader.SequenceComparisonTable;
import com.dfl.report.ExcelReader.SequenceWeldingConditionList;
import com.dfl.report.ExcelReader.WeldPointBoardInformation;
import com.dfl.report.ExcelReader.baseinfoExcelReader;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.commands.print.PropertyTablePrintDialog;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentBOMWindow;
import com.teamcenter.rac.kernel.TCComponentBOMWindowType;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class DealParameterHandler extends AbstractHandler {

	private AbstractAIFUIApplication application;
	private TCSession session;
	private Logger logger = Logger.getLogger(DealParameterHandler.class);
	private ArrayList discrete = new ArrayList();
	private static List<SequenceWeldingConditionList> swc = new ArrayList<SequenceWeldingConditionList>();// 24���к��������趨��
																											// ���к�
	private static List<CurrentandVoltage> cv = new ArrayList<CurrentandVoltage>();// 24���к��������趨�� ������ѹ
	private static List<SFSequenceWeldingConditionList> SFswc = new ArrayList<SFSequenceWeldingConditionList>();// 255���к��������趨��
	private static List<RecommendedPressure> rp = new ArrayList<RecommendedPressure>();// �Ƽ���ѹ��
	private static List<SequenceComparisonTable> sct = new ArrayList<SequenceComparisonTable>();// ���ж��ձ�
	private Map<String, List<String>> MaterialMap;

	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {
		// TODO Auto-generated method stub
		System.out.println("----------------DealParameterHandler------------------");
		logger.info("----------------DealParameterHandler------------------");

		Thread thread = new Thread() {
			public void run() {
				try {
					execute();
				} catch (TCException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		};
		thread.start();
		return null;
	}

	protected void execute() throws TCException {
		// TODO Auto-generated method stub
		application = AIFUtility.getCurrentApplication();
		session = (TCSession) application.getSession();

		InterfaceAIFComponent[] target = application.getTargetComponents();
		// ��ȡѡ��Ķ���
		InterfaceAIFComponent aifComponent = application.getTargetComponent();

		if (aifComponent == null) {
			MessageBox.post("δѡ���κζ���", "��ʾ��Ϣ", MessageBox.ERROR);
			logger.error("δѡ���κζ���");
			return;
		}

		TCComponentBOMLine bomLine = null;

		// �ж�ѡ��Ķ�������
		if (aifComponent instanceof TCComponentBOMLine) {
			bomLine = (TCComponentBOMLine) aifComponent;
		}

		if (bomLine == null) {
			MessageBox.post("��ѡ��BOP�ж���", "��ʾ��Ϣ", MessageBox.ERROR);
			logger.error("��ѡ��BOP�ж���");
			return;
		}
		TCComponentBOMLine topbomLine = bomLine.window().getTopBOMLine();

		if (!topbomLine.getItemRevision().isTypeOf("B8_BIWPlantBOPRevision")) {
			MessageBox.post("��ѡ��BOP�ж���", "��ʾ��Ϣ", MessageBox.ERROR);
			logger.error("��ѡ��BOP�ж���");
			return;
		}

		String type = bomLine.getItemRevision().getType();
		System.out.println("type:" + type);

		dealParameter(bomLine);
	}

	// ���㺸���ϵĲ������ԣ�����ֵ
	private void dealParameter(TCComponentBOMLine bomLine) throws TCException {
		// TODO Auto-generated method stub
		// ��ȡ��ѡ����Note����
		TCPreferenceService ts = session.getPreferenceService();
		if (!ts.isDefinitionExistForPreference("B8_Calculation_Parameter_Name")) {
			MessageBox.post("������ѡ��B8_Calculation_Parameter_Nameδ���壬����ϵϵͳ����Ա��", "��ʾ��Ϣ", MessageBox.ERROR);
			logger.error("������ѡ��B8_Calculation_Parameter_Nameδ����");
			return;
		}
		// ��ȡ���϶��ձ�
		MaterialMap = baseinfoExcelReader.getMaterialComparisonTable(application, "DFL_MaterialMapping");
		if (MaterialMap == null || MaterialMap.size() < 1) {
			System.out.println("δ�ҵ����϶��ձ�");
			MessageBox.post("δ���ö��ձ�DFL_MaterialMapping������ϵϵͳ����Ա��", "��ʾ��Ϣ", MessageBox.ERROR);
			return;
		}
		// ��ʾ�����������
		// ������ʾ���Ȳ����ִ�в���
		ReportViwePanel viewPanel = new ReportViwePanel("���Ӳ�������");
		viewPanel.setVisible(true);

		// ��ȡ����
		TCComponentBOMLine topbl = bomLine.window().getTopBOMLine();

		// ��ȡ�������
		Object[] obj = baseinfoExcelReader.getCalculationParameter(application, "B8_Calculation_Parameter_Name");
		if (obj != null) {
			if (obj[0] != null) {
				swc = (List<SequenceWeldingConditionList>) obj[0];
			} else {
				logger.error("δ��ȡ��24���к��������趨�� ���к���Ϣ��");
				System.out.println("δ��ȡ��24���к��������趨�� ���к���Ϣ��");
				viewPanel.addInfomation("δ��ȡ��24���к��������趨�� ���к���Ϣ��", 100, 100);
				viewPanel.dispose();
				MessageBox.post("δ�ҵ����Ӳ����������", "��ʾ��Ϣ", MessageBox.ERROR);
				return;
			}
			if (obj[1] != null) {
				cv = (List<CurrentandVoltage>) obj[1];
			} else {
				logger.error("δ��ȡ��24���к��������趨�� ������ѹ��Ϣ��");
				System.out.println("δ��ȡ��24���к��������趨�� ������ѹ��Ϣ��");
				viewPanel.addInfomation("δ��ȡ��24���к��������趨�� ������ѹ��Ϣ��", 100, 100);
				viewPanel.dispose();
				MessageBox.post("δ�ҵ����Ӳ����������", "��ʾ��Ϣ", MessageBox.ERROR);
				return;
			}
			if (obj[2] != null) {
				SFswc = (List<SFSequenceWeldingConditionList>) obj[2];
			} else {
				logger.error("δ��ȡ��255���к��������趨����Ϣ��");
				System.out.println("δ��ȡ��255���к��������趨����Ϣ��");
				viewPanel.addInfomation("δ��ȡ��255���к��������趨����Ϣ��", 100, 100);
				viewPanel.dispose();
				MessageBox.post("δ�ҵ����Ӳ����������", "��ʾ��Ϣ", MessageBox.ERROR);
				return;
			}
			if (obj[3] != null) {
				rp = (List<RecommendedPressure>) obj[3];
			} else {
				logger.error("δ��ȡ���Ƽ���ѹ����Ϣ��");
				System.out.println("δ��ȡ���Ƽ���ѹ����Ϣ��");
				viewPanel.addInfomation("δ��ȡ���Ƽ���ѹ����Ϣ��", 100, 100);
				viewPanel.dispose();
				MessageBox.post("δ�ҵ����Ӳ����������", "��ʾ��Ϣ", MessageBox.ERROR);
				return;
			}
			if (obj[4] != null) {
				sct = (List<SequenceComparisonTable>) obj[4];
			} else {
				logger.error("δ��ȡ�����ж��ձ���Ϣ��");
				System.out.println("δ��ȡ�����ж��ձ���Ϣ��");
				viewPanel.addInfomation("δ��ȡ�����ж��ձ���Ϣ��", 100, 100);
				viewPanel.dispose();
				MessageBox.post("δ�ҵ����Ӳ����������", "��ʾ��Ϣ", MessageBox.ERROR);
				return;
			}
		} else {
			viewPanel.addInfomation("δ�ҵ����Ӳ����������", 100, 100);
			logger.error("δ�ҵ����Ӳ����������");
			viewPanel.dispose();
			MessageBox.post("δ�ҵ����Ӳ����������", "��ʾ��Ϣ", MessageBox.ERROR);
			return;
		}
		viewPanel.addInfomation("��ȡ��ز�����Ϣ���...\n", 20, 100);

		List<WeldPointBoardInformation> weldboradlist = getBaseinfomation(topbl, "222.������Ϣ");
		if (weldboradlist == null || weldboradlist.size() < 1) {
			viewPanel.addInfomation("������ϢΪ�գ���ȷ���Ƿ������ɻ�����Ϣ��", 100, 100);
			logger.error("������ϢΪ�գ���ȷ���Ƿ������ɻ�����Ϣ��");
			viewPanel.dispose();
			MessageBox.post("������ϢΪ�գ���ȷ���Ƿ������ɻ�����Ϣ��", "��ʾ��Ϣ", MessageBox.ERROR);
			return;
		}
		viewPanel.addInfomation("��ȡ������Ϣ���\n", 40, 100);
		discrete = getAllDiscrete(bomLine);
		if (discrete == null || discrete.size() < 1) {

			viewPanel.addInfomation("�㺸��������Ϊ�գ���ȷ����ѡ�����´��ڵ㺸�������ݣ�", 100, 100);
			logger.error("�㺸��������Ϊ�գ���ȷ����ѡ�����´��ڵ㺸�������ݣ�");
			viewPanel.dispose();
			MessageBox.post("�㺸��������Ϊ�գ���ȷ����ѡ�����´��ڵ㺸�������ݣ�", "��ʾ��Ϣ", MessageBox.ERROR);
			return;
		}
		viewPanel.addInfomation("��ȡ������Ϣ���...\n", 60, 100);

		// ѭ��������Ϣ�����㲢��ȡ��������ֵ
		Map<TCComponentBOMLine, String[]> paramap = new HashMap<TCComponentBOMLine, String[]>();
		// ����񱡰������ʾ�û�
		ArrayList bzlist = new ArrayList();
		// �����жϺ����ظ��������ظ�ִ��
		ArrayList cflist = new ArrayList();
		// ���ڼ�¼�޺�ǹ�Ĺ�������
		String discretenames = "";
		// ���ڼ�¼�������ŷ����������Եĺ�ǹ����
		String gunnames = "";

		for (int i = 0; i < discrete.size(); i++) {

			TCComponentBOMLine bl = (TCComponentBOMLine) discrete.get(i);
			ArrayList weld = Util.getChildrenByBOMLine(bl, "WeldPointRevision");
			ArrayList gun = Util.getChildrenByBOMLine(bl, "B8_BIWGunRevision");
			String direname = Util.getProperty(bl, "bl_rev_object_name");
			// ֻ��R��ͷ�Ĳ���Ҫ�ж�
			String GunType = "";
			if (direname.length() > 1) {
				if (direname.substring(0, 1).equals("R")) {
					// ֻ�й������к�����ж���û��ǹ
					if (weld != null && weld.size() > 0) {
						if (gun != null && gun.size() > 0) {
							TCComponentBOMLine gbl = (TCComponentBOMLine) gun.get(0);
							GunType = Util.getProperty(gbl.getItemRevision(), "b8_GunType");
							if (GunType.isEmpty()) {
								if (gunnames.isEmpty()) {
									gunnames = direname;
								} else {
									gunnames = gunnames + "," + direname;
								}
							}
						} else {
							if (weld != null && weld.size() > 0) {
								if (discretenames.isEmpty()) {
									discretenames = direname;
								} else {
									discretenames = discretenames + "," + direname;
								}
							}
						}
					}

				}
			}

			if (weld != null && weld.size() > 0) {
				for (int k = 0; k < weld.size(); k++) {
					TCComponentBOMLine wbl = (TCComponentBOMLine) weld.get(k);
					String weldno = Util.getProperty(wbl, "bl_rev_object_name");
					String[] str = new String[13];
					for (int j = 0; j < weldboradlist.size(); j++) {
						WeldPointBoardInformation wbinfo = weldboradlist.get(j);
						String wbNO = wbinfo.getWeldno();
						if (wbNO.equals(weldno)) {
							if (!cflist.contains(weldno)) {
								cflist.add(weldno);
								//��gagi�����ݲ��ʶ��ձ��жϺ����Ƿ���Ҫ�������
								String Partmaterial = wbinfo.getPartmaterial1();								
								String Partmateria2 = wbinfo.getPartmaterial2();								
								String Partmateria3 = wbinfo.getPartmaterial3();
								String gagi1 = wbinfo.getGagi1();								
								String gagi2 = wbinfo.getGagi2();								
								String gagi3 = wbinfo.getGagi3();
								boolean isCalculate = true;
								if(MaterialMap!=null)
								{
									for(Map.Entry<String, List<String>> entry: MaterialMap.entrySet())
									{
										String MaterialNo = entry.getKey();
										List<String> infolist = entry.getValue();
										if(!"GA".equals(gagi1) && !"GI".equals(gagi1))
										{
											if(Util.getIsEqueal(Partmaterial, MaterialNo))
											{
												if("��".equalsIgnoreCase(infolist.get(1)))
												{
													isCalculate = false;
													break;
												}								
											}
										}
										if(!"GA".equals(gagi2) && !"GI".equals(gagi2))
										{
											if(Util.getIsEqueal(Partmateria2, MaterialNo))
											{
												if("��".equalsIgnoreCase(infolist.get(1)))
												{
													isCalculate = false;
													break;
												}								
											}
										}
										if(!"GA".equals(gagi3) && !"GI".equals(gagi3))
										{
											if(Util.getIsEqueal(Partmateria3, MaterialNo))
											{
												if("��".equalsIgnoreCase(infolist.get(1)))
												{
													isCalculate = false;
													break;
												}								
											}
										}																							
									}
								}
								if(!isCalculate)
								{
									str[0] = "";
									str[1] = "";
									str[2] = "";
									str[3] = "";
									str[4] = "";
									str[5] = "";
									str[6] = "";
									str[7] = "";
									str[8] = "";
									str[9] = "";
									str[10] = "";
									str[11] = "";
									str[12] = "";
									paramap.put(wbl, str);
									System.out.println(weldno + "������ڲ���Ҫ���������������");
									break; //�ú�����ڲ�����İ��
								}
								
								String boradnum = wbinfo.getLayersnum();// �����
								String basethickness = wbinfo.getBasethickness();// ��׼���
								String sheetstrength1 = wbinfo.getStrength1();// ���1ǿ��
								String sheetstrength2 = wbinfo.getStrength2();// ���2ǿ��
								String sheetstrength3 = wbinfo.getStrength3();// ���3ǿ��
								String partthickness1 = wbinfo.getPartthickness1();// ���1���
								String partthickness2 = wbinfo.getPartthickness2();// ���2���
								String partthickness3 = wbinfo.getPartthickness3();// ���3���
//								String gagi1 = wbinfo.getGagi1();// ���1GA/GI��
//								String gagi2 = wbinfo.getGagi2();// ���2GA/GI��
//								String gagi3 = wbinfo.getGagi3();// ���3GA/GI��

								// �Ƽ���ѹ������
								String Repressure = "";
								Repressure = getRepressure(basethickness, boradnum, sheetstrength1, sheetstrength2,
										sheetstrength3);
								str[0] = Repressure;
								// System.out.println("����"+weldno+"�Ƽ���ѹ��Ϊ��"+Repressure);
								// 24���к��������趨�� �������к�
								String parameterSerialNo24 = "";
								parameterSerialNo24 = getParameterSerialNo24(basethickness, boradnum, gagi1, gagi2,
										gagi3, sheetstrength1, sheetstrength2, sheetstrength3);

								str[1] = parameterSerialNo24;
								// 255���к��������趨�� �������к�
								// ��ȡ���Ȳ�
								double thicknessdifference = getThicknessDifference(partthickness1, partthickness2,
										partthickness3, boradnum);
								String parameterSerialNo255 = "";
								parameterSerialNo255 = getParameterSerialNo255(basethickness, boradnum, gagi1, gagi2,
										gagi3, sheetstrength1, sheetstrength2, sheetstrength3, thicknessdifference);
								// �����������У��ղ��� ��Ҫ�����������ŷ� ���߼�����
								if (GunType.equals("�ŷ�")) {
									// ֻ��RSW�ŷ������Ƽ����У���Ӧ��
									str[1] = parameterSerialNo255;
									System.out.println("����" + weldno + " ��������(�ղ�)��" + parameterSerialNo255);
									// 255���ж��ձ�
									String SequenceComparison = "";
									SequenceComparison = getSequenceComparison(parameterSerialNo255);
									System.out.println("����" + weldno + " ��������(��Ӧ)��" + parameterSerialNo255);
									str[12] = SequenceComparison;
									str[2] = "";
									str[3] = "";
									str[4] = "";
									str[5] = "";
									str[6] = "";
									str[7] = "";
									str[8] = "";
									str[9] = "";
									str[10] = "";
									str[11] = "";
								} else {
									// ֻ��PSW��RSW������Ҫ�������ֵ
									// 24���к��������趨�� �Ƽ� ����ֵ
									String[] recommendedvalue = getRecommendedvalue(parameterSerialNo24);
									str[2] = recommendedvalue[0];
									str[3] = recommendedvalue[1];
									str[4] = recommendedvalue[2];
									str[5] = recommendedvalue[3];
									str[6] = recommendedvalue[4];
									str[7] = recommendedvalue[5];
									str[8] = recommendedvalue[6];
									str[9] = recommendedvalue[7];
									str[10] = recommendedvalue[8];
									str[11] = recommendedvalue[9];
									str[12] = "";
								}

								String partmaterial1 = wbinfo.getPartmaterial1();
								String partmaterial2 = wbinfo.getPartmaterial2();
								String partmaterial3 = wbinfo.getPartmaterial3();

								String partNos = "";
								// �ж��Ƿ�Ϊ�񱡰�
								boolean flag1 = getJudgingThickSheet(partmaterial1);
								if (flag1) {
									String partNo1 = wbinfo.getPartNo1();
									if (!bzlist.contains(partNo1)) {
										if (partNos.isEmpty()) {
											partNos = partNo1;
										} else {
											partNos = partNos + "," + partNo1;
										}
									}
								}
								boolean flag2 = getJudgingThickSheet(partmaterial2);
								if (flag2) {
									String partNo2 = wbinfo.getPartNo2();
									if (!bzlist.contains(partNo2)) {
										if (partNos.isEmpty()) {
											partNos = partNo2;
										} else {
											partNos = partNos + "," + partNo2;
										}
									}
								}
								boolean flag3 = getJudgingThickSheet(partmaterial3);
								if (flag3) {
									String partNo3 = wbinfo.getPartNo3();
									if (!bzlist.contains(partNo3)) {
										if (partNos.isEmpty()) {
											partNos = partNo3;
										} else {
											partNos = partNos + "," + partNo3;
										}
									}
								}
								if (flag1 || flag2 || flag3) {
									String[] strValues = new String[2];
									strValues[0] = weldno;
									strValues[1] = partNos;
									bzlist.add(strValues);
								} else {
									paramap.put(wbl, str);
								}
								break;
							}
						}
					}
				}
			}

		}
		if (!discretenames.isEmpty()) {
			viewPanel.dispose();
			MessageBox.post("�㺸����" + discretenames + "��û�к�ǹ���޷����㺸�Ӳ�����", "��ʾ��Ϣ", MessageBox.ERROR);
			return;
		}
		if (!gunnames.isEmpty()) {
			viewPanel.dispose();
			MessageBox.post("�㺸����" + gunnames + "��ǹ�������ŷ�/��������Ϊ�գ��޷����㺸�Ӳ�����", "��ʾ��Ϣ", MessageBox.ERROR);
			return;
		}

		viewPanel.addInfomation("���Ӳ���������...\n", 80, 100);

		// �������Բ���ֵ��д����ֵ
		String[] properties = { "bl_WeldPointRevision_b8_RecomWeldForce", "bl_WeldPointRevision_b8_CurrentSerie_Nissan",
				"bl_WeldPointRevision_b8_RiseTime", "bl_WeldPointRevision_b8_CurrentTime1",
				"bl_WeldPointRevision_b8_Current1", "bl_WeldPointRevision_b8_Cool1",
				"bl_WeldPointRevision_b8_CurrentTime2", "bl_WeldPointRevision_b8_Current2",
				"bl_WeldPointRevision_b8_Cool2", "bl_WeldPointRevision_b8_CurrentTime3",
				"bl_WeldPointRevision_b8_Current3", "bl_WeldPointRevision_b8_KeepTime",
				"bl_WeldPointRevision_b8_CurrentSerie_DFL" };

		{
			// ������·
			Util.callByPass(session, true);
		}
		// д����ֵ
		Util.setAllCompsProperty(session, paramap, properties);
		{
			// �ر���·
			Util.callByPass(session, false);
		}
		viewPanel.addInfomation("���Ӳ�������ֵд����ɡ�", 100, 100);

		if (bzlist != null && bzlist.size() > 0) {
			viewPanel.dispose();
			printUnAssigned(bzlist);
		}
	}

	public void printUnAssigned(ArrayList list) {
		if (list != null && !list.isEmpty()) {

			int count = list.size();
			String[] serializable;
			Object[][] arrobject = new Object[count][2];
			int n = 0;
			while (n < count) {
				serializable = (String[]) list.get(n);
				arrobject[n][0] = serializable[0];
				arrobject[n][1] = serializable[1];
				++n;
			}
			Object[] arrobject2 = new String[] { "����ID", "�񱡰��" };
			Frame frame = AIFUtility.getActiveDesktop().getFrame();
			PropertyTablePrintDialog propertyTablePrintDialog = new PropertyTablePrintDialog(frame,
					new JTable(arrobject, arrobject2), "�����嵥");
			propertyTablePrintDialog.setTitle("δ���㺸�Ӳ��������嵥");
			propertyTablePrintDialog.setVisible(true);
			propertyTablePrintDialog.setAlwaysOnTop(true);
		}
	}

	// �ж��Ƿ�Ϊ�񱡰�
	private boolean getJudgingThickSheet(String partmaterial1) {
		// TODO Auto-generated method stub
		boolean flag = false;
		int count1 = 0;
		int count2 = 0;
		String str = "";
		if (partmaterial1 != null) {
			str = partmaterial1;
		}
		count1 = (str.length() - str.replace("SP", "").length()) / "SP".length();
		count2 = (str.length() - str.replace("RP", "").length()) / "RP".length();

		if (count1 + count2 > 1) {
			flag = true;
		}
		return flag;
	}

	// 255���ж��ձ�
	public static String getSequenceComparison(String parameterSerialNo255) {
		// TODO Auto-generated method stub
		String SequenceComparison = "";
		for (int i = 0; i < sct.size(); i++) {
			SequenceComparisonTable sctable = sct.get(i);
			Map<String, String> map = sctable.getValues();
			if (map.containsKey("S" + parameterSerialNo255)) {
				String value = map.get("S" + parameterSerialNo255);
				if (value.trim().length() < 2) {
					value = "0" + value;
				}
				SequenceComparison = sctable.getParameterGroup() + "-" + value;
				break;
			}
		}

		return SequenceComparison;
	}

	// 255���к��������趨�� �������к�
	public static String getParameterSerialNo255(String basethickness, String boradnum, String gagi1, String gagi2,
			String gagi3, String sheetstrength1, String sheetstrength2, String sheetstrength3,
			double thicknessdifference) {
		// TODO Auto-generated method stub
		String parameterSerialNo255 = "";
		int lnum = 0; // �������
		int ganum = 0; // GA������
		int high = 0;// ��ǿ������
		if (gagi1.isEmpty()) {
			lnum++;
		}
		if (gagi2.isEmpty()) {
			lnum++;
		}
		if (gagi3.isEmpty()) {
			lnum++;
		}
		if (boradnum.equals("2")) {
			lnum--;
		}
		if (gagi1.equals("GA") || gagi1.equals("GI")) {
			ganum++;
		}
		if (gagi2.equals("GA") || gagi1.equals("GI")) {
			ganum++;
		}
		if (gagi3.equals("GA") || gagi1.equals("GI")) {
			ganum++;
		}
		if (!sheetstrength1.isEmpty()) {
			high++;
		}
		if (!sheetstrength2.isEmpty()) {
			high++;
		}
		if (!sheetstrength3.isEmpty()) {
			high++;
		}
		for (int i = 0; i < SFswc.size(); i++) {
			SFSequenceWeldingConditionList sfsw = SFswc.get(i);
			String value = sfsw.getBasethickness();
			if (Util.isNumber(value) && Util.isNumber(basethickness)) {
				if (Double.parseDouble(value) == Double.parseDouble(basethickness)) {
					if (boradnum.equals("2")) {
						if (lnum == 2 && high != 0) {
							parameterSerialNo255 = sfsw.getBvalue();
						}
						if (lnum == 2 && high == 0) {
							parameterSerialNo255 = sfsw.getCvalue();
						}
						if (ganum == 1 && thicknessdifference <= 2.4) {
							parameterSerialNo255 = sfsw.getDvalue();
						}
						if (ganum == 1 && thicknessdifference > 2.4) {
							parameterSerialNo255 = sfsw.getEvalue();
						}
						if (ganum == 2 && thicknessdifference <= 2.4) {
							parameterSerialNo255 = sfsw.getFvalue();
						}
						if (ganum == 2 && thicknessdifference > 2.4) {
							parameterSerialNo255 = sfsw.getGvalue();
						}
					}
					if (boradnum.equals("3")) {
						if (lnum == 3 && (high == 2 || high == 3)) {
							parameterSerialNo255 = sfsw.getHvalue();
						}
						if (lnum == 3 && (high == 0 || high == 1)) {
							parameterSerialNo255 = sfsw.getIvalue();
						}
						if (ganum == 1 && thicknessdifference <= 2.4) {
							parameterSerialNo255 = sfsw.getJvalue();
						}
						if (ganum == 1 && thicknessdifference > 2.4) {
							parameterSerialNo255 = sfsw.getKvalue();
						}
						if ((ganum == 2 || ganum == 3) && thicknessdifference <= 2.4) {
							parameterSerialNo255 = sfsw.getLvalue();
						}
						if ((ganum == 2 || ganum == 3) && thicknessdifference > 2.4) {
							parameterSerialNo255 = sfsw.getMvalue();
						}
					}
					break;
				}
			}

		}
		return parameterSerialNo255;
	}

	// ��ȡ���Ȳ�
	public static double getThicknessDifference(String partthickness1, String partthickness2, String partthickness3,
			String boradnum) {
		// TODO Auto-generated method stub
		double pk1;
		double pk2;
		double pk3;
		double thicknessdifference = 0;
		if (Util.isNumber(partthickness1)) {
			pk1 = Double.parseDouble(partthickness1);
		} else {
			pk1 = -1.0;
		}
		if (Util.isNumber(partthickness2)) {
			pk2 = Double.parseDouble(partthickness2);
		} else {
			pk2 = -1.0;
		}
		if (Util.isNumber(partthickness3)) {
			pk3 = Double.parseDouble(partthickness3);
		} else {
			pk3 = -1.0;
		}
		if (boradnum.equals("2")) {
			if (pk1 == -1.0) {
				if (pk2 < pk3) {
					thicknessdifference = pk3 / pk2;
				} else {
					thicknessdifference = pk2 / pk3;
				}
			}
			if (pk2 == -1.0) {
				if (pk1 < pk3) {
					thicknessdifference = pk3 / pk1;
				} else {
					thicknessdifference = pk1 / pk3;
				}
			}
			if (pk3 == -1.0) {
				if (pk1 < pk2) {
					thicknessdifference = pk2 / pk1;
				} else {
					thicknessdifference = pk1 / pk2;
				}
			}
		}
		if (boradnum.equals("3")) {
			String minstr = getMinnum(partthickness1, partthickness2, partthickness3);
			String maxstr = getMaxnum(partthickness1, partthickness2, partthickness3);
			double min = Double.parseDouble(minstr);
			double max = Double.parseDouble(maxstr);
			thicknessdifference = max / min;
		}
		BigDecimal bd = new BigDecimal(thicknessdifference);
		BigDecimal fact = bd.setScale(1, BigDecimal.ROUND_HALF_UP);
		thicknessdifference = fact.doubleValue();
		return thicknessdifference;
	}

	// 24���к��������趨�� �Ƽ� ����ֵ
	public static String[] getRecommendedvalue(String parameterSerialNo24) {
		// TODO Auto-generated method stub
		String[] recommendedvalue = new String[10];
		for (int i = 0; i < cv.size(); i++) {
			CurrentandVoltage cvotage = cv.get(i);
			String serialNo = cvotage.getSequenceNo();
			if (serialNo != null && serialNo.equals(parameterSerialNo24)) {
				recommendedvalue[0] = cvotage.getBvalue();// ����ʱ��
				recommendedvalue[1] = cvotage.getCvalue();// ��һ ͨ��ʱ��
				recommendedvalue[2] = cvotage.getEvalue();// ��һ ͨ�����
				recommendedvalue[3] = cvotage.getFvalue();// ��ȴʱ��һ
				recommendedvalue[4] = cvotage.getGvalue();// �ڶ�ͨ��ʱ��
				recommendedvalue[5] = cvotage.getIvalue();// �ڶ�ͨ�����
				recommendedvalue[6] = cvotage.getJvalue();// ��ȴʱ���
				recommendedvalue[7] = cvotage.getKvalue();// ���� ͨ��ʱ��
				recommendedvalue[8] = cvotage.getMvalue();// ���� ͨ�����
				recommendedvalue[9] = cvotage.getNvalue();// ����
				break;
			}
		}
		return recommendedvalue;
	}

	// 24���к��������趨�� �������к�
	public static String getParameterSerialNo24(String basethickness, String boradnum, String gagi1, String gagi2,
			String gagi3, String sheetstrength1, String sheetstrength2, String sheetstrength3) {
		// TODO Auto-generated method stub
		String parameterSerialNo24 = "";
		int lnum = 0; // �������
		int ginum = 0; // GI������
		int ganum = 0; // GA������
		int high = 0;// ��ǿ������

		if (gagi1.isEmpty()) {
			lnum++;
		}
		if (gagi2.isEmpty()) {
			lnum++;
		}
		if (gagi3.isEmpty()) {
			lnum++;
		}
		if (boradnum.equals("2")) {
			lnum--;
		}

		if (gagi1.equals("GA")) {
			ganum++;
		}
		if (gagi2.equals("GA")) {
			ganum++;
		}
		if (gagi3.equals("GA")) {
			ganum++;
		}
		if (gagi1.equals("GI")) {
			ginum++;
		}
		if (gagi2.equals("GI")) {
			ginum++;
		}
		if (gagi3.equals("GI")) {
			ginum++;
		}
		if (!sheetstrength1.isEmpty()) {
			high++;
		}
		if (!sheetstrength2.isEmpty()) {
			high++;
		}
		if (!sheetstrength3.isEmpty()) {
			high++;
		}
		for (int i = 0; i < swc.size(); i++) {
			SequenceWeldingConditionList swcl = swc.get(i);
			String thick = swcl.getBasethickness();
			if (Util.isNumber(thick) && Util.isNumber(basethickness)) {
				if (Double.parseDouble(thick) == Double.parseDouble(basethickness)) {
					// ��������м���GA������GI��ʱ����GA�ĵ���GI�������ǡ�

					if (boradnum.equals("2")) {
						if (lnum == 2 && high == 0) {
							parameterSerialNo24 = swcl.getBvalue();
						}
						if (lnum == 2 && high != 0) {
							parameterSerialNo24 = swcl.getCvalue();
						}
						if (lnum == 1 && ganum == 1 && high == 0) {
							parameterSerialNo24 = swcl.getDvalue();
						}
						if (lnum == 1 && ganum == 1 && high != 0) {
							parameterSerialNo24 = swcl.getEvalue();
						}
						if (lnum == 1 && ganum == 1 && high != 0) {
							parameterSerialNo24 = swcl.getEvalue();
						}
						if (ganum == 2 && high == 0) {
							parameterSerialNo24 = swcl.getFvalue();
						}
						if (ganum == 2 && high != 0) {
							parameterSerialNo24 = swcl.getGvalue();
						}
						if (ginum == 1 && ganum == 1 && high == 0) {
							parameterSerialNo24 = swcl.getHvalue();
						}
						if (ginum == 1 && ganum == 1 && high != 0) {
							parameterSerialNo24 = swcl.getIvalue();
						}
						if (ginum == 2 && high == 0) {
							parameterSerialNo24 = swcl.getJvalue();
						}
						if (ginum == 2 && high != 0) {
							parameterSerialNo24 = swcl.getKvalue();
						}
					}
					if (boradnum.equals("3")) {
						if (lnum == 3 && high == 0) {
							parameterSerialNo24 = swcl.getLvalue();
						}
						if (lnum == 3 && high != 0) {
							parameterSerialNo24 = swcl.getMvalue();
						}
						if (ganum == 1 && ginum == 0 && high == 0) {
							parameterSerialNo24 = swcl.getNvalue();
						}
						if (ganum == 1 && ginum == 0 && high != 0) {
							parameterSerialNo24 = swcl.getOvalue();
						}
						if (ganum == 2 && ginum == 0 && high == 0) {
							parameterSerialNo24 = swcl.getPvalue();
						}
						if (ganum == 2 && ginum == 0 && high != 0) {
							parameterSerialNo24 = swcl.getQvalue();
						}
						if (ganum == 3 && ginum == 0 && high == 0) {
							parameterSerialNo24 = swcl.getRvalue();
						}
						if (ganum == 3 && ginum == 0 && high != 0 && high != 3) {
							parameterSerialNo24 = swcl.getSvalue();
						}
						if (ganum == 3 && ginum == 0 && high == 3) {
							parameterSerialNo24 = swcl.getTvalue();
						}
						if (ganum == 0 && ginum == 1 && high == 0) {
							parameterSerialNo24 = swcl.getUvalue();
						}
						if (ganum == 0 && ginum == 1 && high != 0) {
							parameterSerialNo24 = swcl.getVvalue();
						}
						if (lnum == 1 && ganum != 2 && high == 0) {
							parameterSerialNo24 = swcl.getWvalue();
						}
						if (lnum == 1 && ganum != 2 && high != 0) {
							parameterSerialNo24 = swcl.getXvalue();
						}
						if (lnum == 0 && ganum != 3 && high == 0) {
							parameterSerialNo24 = swcl.getYvalue();
						}
						if (lnum == 0 && ganum != 3 && high != 0 && high != 3) {
							parameterSerialNo24 = swcl.getZvalue();
						}
						if (lnum == 0 && ganum != 3 && high == 3) {
							parameterSerialNo24 = swcl.getAAvalue();
						}
					}
					break;
				}
			}

		}

		return parameterSerialNo24;
	}

	// ��ȡ��ѹ��
	public static String getRepressure(String basethickness, String boradnum, String sheetstrength1,
			String sheetstrength2, String sheetstrength3) {
		// TODO Auto-generated method stub
		String repressure = "";
		String distinguish = ""; // ����
		int num1 = 0;// 440��������
		int num2 = 0;// 440
		int num3 = 0;// 590Mpa780Mpa980Mpa
		// �������1180ǿ�Ȱ�ģ�����������У�Ĭ��Ϊ��
		int shstrength1 = getInteger(sheetstrength1);
		int shstrength2 = getInteger(sheetstrength2);
		int shstrength3 = getInteger(sheetstrength3);
		if (shstrength1 == 1180 || shstrength2 == 1180 || shstrength3 == 1180) {
			repressure = "";
			return repressure;
		}
		// �������1350ǿ�Ȱ�ģ�����ΪI
		if (shstrength1 == 1350 || shstrength2 == 1350 || shstrength3 == 1350) {
			distinguish = "��";
		} else {
			if (sheetstrength1.isEmpty()) {
				num1++;
			}
			if (shstrength1 == 440) {
				num2++;
			}
			if (shstrength1 == 590 || shstrength1 == 780 || shstrength1 == 980) {
				num3++;
			}
			if (sheetstrength2.isEmpty()) {
				num1++;
			}
			if (shstrength2 == 440) {
				num2++;
			}
			if (shstrength2 == 590 || shstrength2 == 780 || shstrength2 == 980) {
				num3++;
			}
			if (sheetstrength3.isEmpty()) {
				num1++;
			}
			if (shstrength3 == 440) {
				num2++;
			}
			if (shstrength3 == 590 || shstrength3 == 780 || shstrength3 == 980) {
				num3++;
			}

			// �ȸ����������򣬻�ȡ����,�ٸ���3���
			if (boradnum.equals("2")) {
				if (num1 == 3) {
					distinguish = "��";
				}
				if (num3 == 0 && num2 == 1) {
					distinguish = "��";
				}
				if (num2 == 0 && num3 == 1) {
					distinguish = "��";
				}
				if (num2 == 2) {
					distinguish = "��";
				}
				if (num2 == 1 && num3 == 1) {
					distinguish = "��";
				}
				if (num3 == 2) {
					distinguish = "��";
				}
			} else if (boradnum.equals("3")) {
				if (num1 == 3) {
					distinguish = "��";
				}
				if (num1 == 2 && num2 == 1) {
					distinguish = "��";
				}
				if (num1 == 2 && num3 == 1) {
					distinguish = "��";
				}
				if (num1 == 1 && num2 == 2) {
					distinguish = "��";
				}
				if (num1 == 1 && num2 == 1 && num3 == 1) {
					distinguish = "��";
				}
				if (num1 == 1 && num3 == 2) {
					distinguish = "��";
				}
				if (num2 == 3) {
					distinguish = "��";
				}
				if (num2 == 2 && num3 == 1) {
					distinguish = "��";
				}
				if (num2 == 1 && num3 == 2) {
					distinguish = "��";
				}
				if (num3 == 3) {
					distinguish = "��";
				}
			} else {
				repressure = "";
				return repressure;
			}

		}
		for (int i = 0; i < rp.size(); i++) {
			RecommendedPressure repre = rp.get(i);
			String thickness = repre.getBasethickness();
			if (Util.isNumber(thickness) && Util.isNumber(basethickness)) {
				if (Double.parseDouble(thickness) == Double.parseDouble(basethickness)) {
					if (distinguish.equals("��")) {
						repressure = repre.getBvalue();
					}
					if (distinguish.equals("��")) {
						repressure = repre.getCvalue();
					}
					if (distinguish.equals("��")) {
						repressure = repre.getDvalue();
					}
					if (distinguish.equals("��")) {
						repressure = repre.getEvalue();
					}
					if (distinguish.equals("��")) {
						repressure = repre.getFvalue();
					}
					break;
				}
			}

		}

		return repressure;
	}

	// ��ȡBOP�»�����Ϣ�к��������Ϣ
	public static List<WeldPointBoardInformation> getBaseinfomation(TCComponentBOMLine topbl, String procName) {
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

	// ��ȡBOP�е����к���
	public ArrayList<TCComponentBOMLine> getAllDiscrete(TCComponentBOMLine topbl) {
		// TODO Auto-generated method stub

		ArrayList<TCComponentBOMLine> welds = new ArrayList<TCComponentBOMLine>();
		System.out.println("���ڲ�������Ӧ����");

		ArrayList qclist = new ArrayList();
		// ����BOP��ѯ���еĵ㺸����
		String typename = Util.getObjectDisplayName(session, "B8_BIWDiscreteOP");
		String[] propertys = new String[] { "bl_item_object_type", "bl_item_object_type" };
		String[] values = new String[] { typename, typename };

		ArrayList partList = Util.searchBOMLine(topbl, "OR", propertys, "==", values);
		System.out.println("�㺸���򼯺ϣ�" + partList.toString());

		return partList;
	}

	/*
	 * ȡ��Сֵ
	 */
	public static String getMinnum(String str1, String str2, String str3) {
		String minstr = "";
		if (str1 == null || str1.isEmpty()) {
			str1 = "9999";
		}
		if (str2 == null || str2.isEmpty()) {
			str2 = "9999";
		}
		if (str3 == null || str3.isEmpty()) {
			str3 = "9999";
		}
		if (Double.parseDouble(str1) > Double.parseDouble(str2)) {
			if (Double.parseDouble(str2) > Double.parseDouble(str3)) {
				minstr = str3;
			} else {
				minstr = str2;
			}
		} else {
			if (Double.parseDouble(str1) > Double.parseDouble(str3)) {
				minstr = str3;
			} else {
				minstr = str1;
			}
		}
//		if (minstr.equals("9999")) {
//			minstr = "";
//		}
		return minstr;
	}

	/*
	 * ȡ���ֵ
	 */
	public static String getMaxnum(String str1, String str2, String str3) {
		String maxstr = "";
		if (str1 == null || str1.isEmpty()) {
			str1 = "-1";
		}
		if (str2 == null || str2.isEmpty()) {
			str2 = "-1";
		}
		if (str3 == null || str3.isEmpty()) {
			str3 = "-1";
		}
		if (Double.parseDouble(str1) > Double.parseDouble(str2)) {
			if (Double.parseDouble(str1) > Double.parseDouble(str3)) {
				maxstr = str1;
			} else {
				maxstr = str3;
			}
		} else {
			if (Double.parseDouble(str2) > Double.parseDouble(str3)) {
				maxstr = str2;
			} else {
				maxstr = str3;
			}
		}
//		if (maxstr.equals("-1")) {
//			maxstr = "";
//		}
		return maxstr;
	}

	/*
	 * �ַ�ת��������
	 */
	public static int getInteger(String str) {
		int num = -1;
		if (Util.isNumber(str)) {
			num = (int) Double.parseDouble(str);
		}
		return num;
	}

}
