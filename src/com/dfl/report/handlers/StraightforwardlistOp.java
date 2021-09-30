package com.dfl.report.handlers;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.logging.Logger;

import javax.swing.SwingUtilities;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dfl.report.ExcelReader.CoverInfomation;
import com.dfl.report.ExcelReader.baseinfoExcelReader;
import com.dfl.report.util.FileUtil;
import com.dfl.report.util.NewOutputDataToExcel;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentDatasetType;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentTcFile;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class StraightforwardlistOp {

	private AbstractAIFUIApplication app;
	private ReportViwePanel viewPanel;
	// ��С������
	private HashMap<String, String> rule = new HashMap<String, String>();
	SimpleDateFormat df = new SimpleDateFormat("yyyyMMdd  HH");// �������ڸ�ʽ
	private TCComponent folder;
	private static Logger logger = Logger.getLogger(baseinfoExcelReader.class.getName()); // ��־��ӡ��
	private InterfaceAIFComponent[] ifc;
	private TCSession session;
	private InputStream inputStream;

	public StraightforwardlistOp(AbstractAIFUIApplication app, TCComponent savefolder, InterfaceAIFComponent[] ifc,
			TCSession session, InputStream inputStream) {
		// TODO Auto-generated constructor stub
		this.app = app;
		this.folder = savefolder;
		this.ifc = ifc;
		this.session = session;
		this.inputStream = inputStream;
		initUI();
	}

	// ����
	int hsnum = 0;// ��������
	int rswnum = 0;// RSW����

	private void initUI() {
		// TODO Auto-generated method stub
		try {

			// InterfaceAIFComponent[] aifc = app.getTargetComponents();

			TCComponentBOMLine aifbl = (TCComponentBOMLine) ifc[0];
			TCComponentBOMLine topbl = aifbl.window().getTopBOMLine();

			String familiycode = Util.getProperty(topbl.getItemRevision(), "project_ids");
			String vecile = Util.getDFLProjectIdVehicle(familiycode);
			if(vecile==null || vecile.isEmpty()) {
				vecile = familiycode;
			}

//			// ���ݶ���BOP��ѯ���еĺ�װ����
//			String[] propertys = new String[] { "bl_item_object_type", "bl_item_object_type" };
//			String[] values = new String[] { "��װ���߹���", "BIW Process Line" };
//
//			// ���ݺ�װ���߲�ѯ���еĵ㺸����
//			String[] propertys2 = new String[] { "bl_item_object_type", "bl_item_object_type" };
//			String[] values2 = new String[] { "��װ��λ����", "BIW Process Station" };

			viewPanel = new ReportViwePanel("���ɱ���");
			viewPanel.setVisible(true);

			// viewPanel.addInfomation("��ʼ�������...\n", 5, 100);
			viewPanel.addInfomation("��ʼ�������......\n", 10, 100);

			
			// �Ȼ�ȡ���е�������
			List CList = Util.getChildrenByParent(ifc);

			if (CList == null) {
				viewPanel.dispose();
				MessageBox.post("���󣺵�ǰ���͵�PlantBOP�£�û�е㺸�������ݣ�", "��ܰ��ʾ", MessageBox.INFORMATION);				
				return;
			}
			// ��METAL���߷ŵ�������
			ArrayList SortList = getOrderList(CList);

			// �ٸ������������߻�ȡʵ�ʲ���
			// ArrayList partList = getFactLineByParent(SortList);

			if (SortList == null) {
				viewPanel.dispose();
				MessageBox.post("���󣺵�ǰ���͵�PlantBOP�£�û�е㺸�������ݣ�", "��ܰ��ʾ", MessageBox.INFORMATION);
				return;
			}

			viewPanel.addInfomation("", 20, 100);

			XSSFWorkbook book = NewOutputDataToExcel.creatXSSFWorkbook(inputStream);

			// �����е��������
			int total_rownum = 11;// ��ȡģ��ĳ�ʼλ��

			// ��ȡ��С������
			rule = getSizeRule();

			viewPanel.addInfomation("��ʼд���ݣ������ĵȴ�...\n", 40, 100);
			for (int i = 0; i < SortList.size(); i++) {

				viewPanel.addInfomation("", 40, 100);

				TCComponentBOMLine topline = (TCComponentBOMLine) SortList.get(i);
				// ��ȡ��λ����
				ArrayList discreteList = getStateChildrenByParent(topline);

				/*
				 * ***************************** ��������ߵ�����ȥ�ж��Ƿ�Ϊ���ֻ����������˹���
				 */
				int ajm = getIsAJ(topline);
				// ���������� ��Ҫƴ������ʽ����� 20191202�޸�				
				String plinename = topline.getProperty("bl_rev_object_name");
				long startTime = System.currentTimeMillis(); // ��ȡ��ʼʱ��
				ArrayList list = getArrayListData(discreteList, ajm);
				long endTime = System.currentTimeMillis(); // ��ȡ����ʱ��
				System.out.println("��ȡһ���������ݣ� " + (endTime - startTime) + "ms");

				list.add(plinename);

				System.out.println("�������ƣ�" + list.get(list.size() - 1) + "/�����µ�����������" + list.get(list.size() - 2));

				NewOutputDataToExcel.writeDataToSheet(book, list, hsnum, rswnum, total_rownum, viewPanel, ajm);

				if (total_rownum == 11) {
					total_rownum = total_rownum + (int) (list.get(list.size() - 2)) - 1;
				} else {
					total_rownum = total_rownum + (int) (list.get(list.size() - 2));
				}

				System.out.println("��������" + total_rownum);
			}
			viewPanel.addInfomation("", 60, 100);
			// ���ɾ��ģ���в����û����еĹ�ʽ
			NewOutputDataToExcel.dealTotalRowFormula(book, viewPanel);

			String date = df.format(new Date());
			String datasetname = vecile + "ֱ�ͼ����" + "_" + date + "ʱ";
			String filename = Util.formatString(datasetname);

			NewOutputDataToExcel.exportFile(book, filename);

			viewPanel.addInfomation("", 80, 100);

			// String fullFileName = FileUtil.getReportFileName("ֱ���嵥��");

			Util.saveFiles(filename, datasetname, folder, session, "AD");
			// NewOutputDataToExcel.openFile(fullFileName);

			viewPanel.addInfomation("���������ɣ�����ѡ�񱣴���ļ����²鿴��\n", 100, 100);

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private ArrayList getStateChildrenByParent(TCComponentBOMLine topline) {
		// TODO Auto-generated method stub
		ArrayList list = new ArrayList();
		try {
			AIFComponentContext[] childrens = topline.getChildren();
			for (AIFComponentContext chil : childrens) {
				TCComponentBOMLine bl = (TCComponentBOMLine) chil.getComponent();
				AIFComponentContext[] childrens2 = bl.getChildren();
				for (AIFComponentContext chil2 : childrens2) {
					TCComponentBOMLine dbl = (TCComponentBOMLine) chil2.getComponent();
					if(dbl.getItemRevision().isTypeOf("B8_BIWMEProcStatRevision")) {
						list.add(dbl);
					}				
				}
			}
			return list;

		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return list;
	}

	private ArrayList getFactLineByParent(ArrayList cList) throws TCException {
		// TODO Auto-generated method stub
		ArrayList list = new ArrayList();
		for (int i = 0; i < cList.size(); i++) {
			TCComponentBOMLine pbl = (TCComponentBOMLine) cList.get(i);
			AIFComponentContext[] childrens = pbl.getChildren();
			for (AIFComponentContext aif : childrens) {
				TCComponentBOMLine bl = (TCComponentBOMLine) aif.getComponent();
				TCComponentItemRevision rev = bl.getItemRevision();
				if (rev.isTypeOf("B8_BIWMEProcLineRevision")) {
					list.add(bl);
				}
			}
		}
		return list;
	}

	// ��ȡ�����µĵ㺸����
	private ArrayList getArrayListData(ArrayList partList, int ajm) {

		ArrayList list = new ArrayList();// �����ݼ���
		try {

			String PARTS_NAME;// PARTS NAME(OPERATION NAME)��ֵ

			int GUN;// GUN QTY��ֵ
			int RSW;// RSW(MSW)��ֵ
			int PSW;// PSW PTS��ֵ
			int rownum = 0;// ��������Ҫ�����ƶ�������
			/*
			 * ���ݺ�װ��λ�����Ƿ��ж���㺸���������PARTS NAME(OPERATION NAME)
			 * ����ȡֵΪ��װ��λ���գ��㺸�������ƣ�������ֻȡֵ��װ��λ���գ����JR��ͷ�Ĺ�λ���գ����
			 * �ж���㺸���򣬷ֵ㺸�������������Ϣ������Ϊ�㺸�������ơ�������˹����ӹ�λ��AJ��ͷ���������ֹ��򣬺ϲ����
			 */
			for (int i = 0; i < partList.size(); i++) {
				viewPanel.addInfomation("", 40, 100);

				ArrayList ajlist = new ArrayList();// �㺸���򼯺�

				TCComponentBOMLine bl = (TCComponentBOMLine) partList.get(i);
				String plinename = Util.getProperty(bl.parent().getItemRevision(), "b8_LineType")
						+ bl.parent().getProperty("bl_rev_object_name");

				boolean flag = getIsDiscretes(bl); // �ж��Ƿ���ں�����R��ͷ�Ļ����˹���
				// �����಻��Ҫ���ֻ����˹�������λͳ�����
				if (ajm == 3) {
					int LARGE_PARTS = 0;// LARGE PARTS��ֵ
					int SUPERLARGE_PARTS = 0;// SUPERLARGE_PARTS��ֵ
					int MID_PARTS = 0;// MID PARTS��ֵ
					int SMALL_PARTS = 0;// SMALL PARTS��ֵ
					int CLAMP = 0;// CLAMP& UN-CLAMP��ֵ

					String[] str = new String[10];// ��10������
					PARTS_NAME = plinename + " " + bl.getProperty("bl_rev_object_name");

					ajlist = Util.getChildrenByParent(bl);

					str[0] = Integer.toString(rownum + 1);// ���
					str[1] = PARTS_NAME;

					// ���ݺ�װ��λ��ѯ���
					ArrayList part = new ArrayList();// �������

					getSolItmPart(bl, part);

					for (int j = 0; j < part.size(); j++) {
						TCComponentItemRevision rev = (TCComponentItemRevision) part.get(j);
						String partno = Util.getProperty(rev, "dfl9_part_no");
						if (partno.length() > 5) {
							partno = partno.substring(0, 5);
						}
						if (rule.containsKey(partno)) {
							String type = rule.get(partno);
							if (type.equals("SUPER LARGE PARTS")) {
								SUPERLARGE_PARTS++;
							}
							if (type.equals("LARGE PARTS")) {
								LARGE_PARTS++;
							}
							if (type.equals("MID PARTS")) {
								MID_PARTS++;
							}
							if (type.equals("SMALL PARTS")) {
								SMALL_PARTS++;
							}
						}
					}

					if (ajlist != null) {

						// SUPERLARGE_PARTS = 0;
						if (SUPERLARGE_PARTS != 0) {
							str[2] = Integer.toString(SUPERLARGE_PARTS);
						}
						// LARGE_PARTS = 0;
						if (LARGE_PARTS != 0) {
							str[3] = Integer.toString(LARGE_PARTS);
						}
						// MID_PARTS = 0;
						if (MID_PARTS != 0) {
							str[4] = Integer.toString(MID_PARTS);
						}
						// SMALL_PARTS = 0;
						if (SMALL_PARTS != 0) {
							str[5] = Integer.toString(SMALL_PARTS);
						}
						if (CLAMP != 0) {
							str[6] = Integer.toString(CLAMP);
						}
						// CLAMP = 0;

						// ���ݺ�װ��λȥͳ��ǹ�ͺ���
//						ArrayList QList = Util.searchBOMLine(bl, "OR",
//								new String[] { "bl_item_object_type", "bl_item_object_type" }, "==",
//								new String[] { "ǹ", "BIW Gun" });
//						ArrayList QList = getGunNums(bl);						
//						ArrayList HDList = Util.searchBOMLine(bl, "OR",
//								new String[] { "bl_item_object_type", "bl_item_object_type" }, "==",
//								new String[] { "����", "Weld Point" });
						Integer[] hdnums = getWeldPointNums(bl);

//						if (QList != null) {
//							GUN = QList.size();
//							str[7] = Integer.toString(GUN);
//						} else {
//							str[7] = "0";
//						}
						str[7] = getGunNums(bl).toString();
						if (hdnums != null) {
							PSW = hdnums[0];
							RSW = hdnums[1];
							if (RSW != 0) {
								str[8] = hdnums[1].toString();
							}
							if (PSW != 0) {
								str[9] = hdnums[0].toString();
							}
							hsnum = hsnum + PSW + RSW;
							rswnum = rswnum + RSW;
						}
//						if (HDList != null) {
//							PSW = HDList.size();
//							str[9] = Integer.toString(PSW);
//							hsnum = hsnum + PSW;
//						} else {
//							PSW = 0;
//							str[9] = Integer.toString(PSW);
//						}
					}
					rownum++;
					list.add(str);

				}
				// ��Ҫ���ֻ����˹����ߣ�����Ҫע�����Ϲ���͵㺸�����������һ������Ҫ�ϲ�һ�����
				if (ajm == 1) {
					// ��λ���л����˹���
					if (flag) {
						ajlist = Util.getChildrenByBOMLine(bl, "B8_BIWDiscreteOPRevision");// ��ȡ��λ�µĵ㺸����
						ArrayList oplist = Util.getChildrenByBOMLine(bl, "B8_BIWOperationRevision");// ��ȡ��λ�µ����Ϲ���
						for (int j = 0; j < ajlist.size(); j++) {
							int LARGE_PARTS = 0;// LARGE PARTS��ֵ
							int SUPERLARGE_PARTS = 0;// SUPERLARGE_PARTS��ֵ
							int MID_PARTS = 0;// MID PARTS��ֵ
							int SMALL_PARTS = 0;// SMALL PARTS��ֵ
							int CLAMP = 0;// CLAMP& UN-CLAMP��ֵ

							String[] str = new String[10];// ��10������
							TCComponentBOMLine chil = (TCComponentBOMLine) ajlist.get(j);
							String discretename = chil.getProperty("bl_rev_object_name");
							TCComponentBOMLine relatively = null; // ��Ӧ�����Ϲ���
							// ����㺸�������ƺ��ϼ���������һ�£��ϲ����
							if (oplist != null) {
								for (int m = 0; m < oplist.size(); m++) {
									TCComponentBOMLine ch = (TCComponentBOMLine) oplist.get(m);
									String opname = ch.getProperty("bl_rev_object_name");
									if (discretename.equals(opname)) {
										relatively = ch;
										oplist.remove(m);
										break;
									}
								}
							}

							// ���ݺ�װ��λ��ѯ���
							ArrayList part = new ArrayList();// �������

							if (relatively != null) {
								getSolItmPart(relatively, part);
							}

							for (int k = 0; k < part.size(); k++) {
								TCComponentItemRevision rev = (TCComponentItemRevision) part.get(k);
								String partno = Util.getProperty(rev, "dfl9_part_no");
								if (partno.length() > 5) {
									partno = partno.substring(0, 5);
								}
								if (rule.containsKey(partno)) {
									String type = rule.get(partno);
									if (type.equals("SUPER LARGE PARTS")) {
										SUPERLARGE_PARTS++;
									}
									if (type.equals("LARGE PARTS")) {
										LARGE_PARTS++;
									}
									if (type.equals("MID PARTS")) {
										MID_PARTS++;
									}
									if (type.equals("SMALL PARTS")) {
										SMALL_PARTS++;
									}
								}
							}

							PARTS_NAME = plinename + " " + bl.getProperty("bl_rev_object_name") + "(" + discretename
									+ ")";

							str[0] = Integer.toString(rownum + 1);// ���

							str[1] = PARTS_NAME;

							// SUPERLARGE_PARTS = 0;
							if (SUPERLARGE_PARTS != 0) {
								str[2] = Integer.toString(SUPERLARGE_PARTS);
							}
							// LARGE_PARTS = 0;
							if (LARGE_PARTS != 0) {
								str[3] = Integer.toString(LARGE_PARTS);
							}
							// MID_PARTS = 0;
							if (MID_PARTS != 0) {
								str[4] = Integer.toString(MID_PARTS);
							}
							// SMALL_PARTS = 0;
							if (SMALL_PARTS != 0) {
								str[5] = Integer.toString(SMALL_PARTS);
							}
							if (CLAMP != 0) {
								str[6] = Integer.toString(CLAMP);
							}

							ArrayList QList = Util.getChildrenByBOMLine(chil, "B8_BIWGunRevision");
							ArrayList HDList = Util.getChildrenByBOMLine(chil, "WeldPointRevision");
							// ���ݵ㺸����������Ƿ�ΪR��ͷ��ȷ�����������˹�
							if (discretename.substring(0, 1).equals("R")) {
								RSW = HDList.size();
								str[8] = Integer.toString(RSW);
								hsnum = hsnum + RSW;
								rswnum = rswnum + RSW;
							} else {
								if (QList != null) {
									str[7] = Integer.toString(QList.size());
								} else {
									str[7] = "0";
								}
								PSW = HDList.size();
								str[9] = Integer.toString(PSW);
								hsnum = hsnum + PSW;
							}
//							if (HDList != null) {
//								RSW = HDList.size();
//								str[8] = Integer.toString(RSW);
//								hsnum = hsnum + RSW;
//								rswnum = rswnum + RSW;
//							} else {
//								RSW = 0;
//								str[7] = Integer.toString(RSW);
//							}
							rownum++;

							list.add(str);
						}
						// ������Ϲ���û�ж�Ӧ�ĵ㺸����Ҳ��Ҫ���������Ҫͳ�ƺ���
						if (oplist != null && oplist.size() > 0) {
							for (int n = 0; n < oplist.size(); n++) {
								int LARGE_PARTS = 0;// LARGE PARTS��ֵ
								int SUPERLARGE_PARTS = 0;// SUPERLARGE_PARTS��ֵ
								int MID_PARTS = 0;// MID PARTS��ֵ
								int SMALL_PARTS = 0;// SMALL PARTS��ֵ
								int CLAMP = 0;// CLAMP& UN-CLAMP��ֵ

								String[] str = new String[10];// ��10������
								TCComponentBOMLine chil = (TCComponentBOMLine) oplist.get(n);
								String discretename = chil.getProperty("bl_rev_object_name");

								// ���ݺ�װ��λ��ѯ���
								ArrayList part = new ArrayList();// �������

								getSolItmPart(chil, part);

								for (int k = 0; k < part.size(); k++) {
									TCComponentItemRevision rev = (TCComponentItemRevision) part.get(k);
									String partno = Util.getProperty(rev, "dfl9_part_no");
									if (partno.length() > 5) {
										partno = partno.substring(0, 5);
									}
									if (rule.containsKey(partno)) {
										String type = rule.get(partno);
										if (type.equals("SUPER LARGE PARTS")) {
											SUPERLARGE_PARTS++;
										}
										if (type.equals("LARGE PARTS")) {
											LARGE_PARTS++;
										}
										if (type.equals("MID PARTS")) {
											MID_PARTS++;
										}
										if (type.equals("SMALL PARTS")) {
											SMALL_PARTS++;
										}
									}
								}

								PARTS_NAME = plinename + " " + bl.getProperty("bl_rev_object_name") + "(" + discretename
										+ ")";

								str[0] = Integer.toString(rownum + 1);// ���

								str[1] = PARTS_NAME;

								// SUPERLARGE_PARTS = 0;
								if (SUPERLARGE_PARTS != 0) {
									str[2] = Integer.toString(SUPERLARGE_PARTS);
								}
								// LARGE_PARTS = 0;
								if (LARGE_PARTS != 0) {
									str[3] = Integer.toString(LARGE_PARTS);
								}
								// MID_PARTS = 0;
								if (MID_PARTS != 0) {
									str[4] = Integer.toString(MID_PARTS);
								}
								// SMALL_PARTS = 0;
								if (SMALL_PARTS != 0) {
									str[5] = Integer.toString(SMALL_PARTS);
								}
								if (CLAMP != 0) {
									str[6] = Integer.toString(CLAMP);
								}
								list.add(str);
							}
						}
					}
					// ��λ���޻����˹���
					else {
						int LARGE_PARTS = 0;// LARGE PARTS��ֵ
						int SUPERLARGE_PARTS = 0;// SUPERLARGE_PARTS��ֵ
						int MID_PARTS = 0;// MID PARTS��ֵ
						int SMALL_PARTS = 0;// SMALL PARTS��ֵ
						int CLAMP = 0;// CLAMP& UN-CLAMP��ֵ

						PARTS_NAME = plinename + " " + bl.getProperty("bl_rev_object_name");

						// ���ݺ�װ��λ��ѯ���
						ArrayList part = new ArrayList();// �������

						getSolItmPart(bl, part);

						for (int j = 0; j < part.size(); j++) {
							TCComponentItemRevision rev = (TCComponentItemRevision) part.get(j);
							String partno = Util.getProperty(rev, "dfl9_part_no");
							if (partno.length() > 5) {
								partno = partno.substring(0, 5);
							}
							if (rule.containsKey(partno)) {
								String type = rule.get(partno);
								if (type.equals("SUPER LARGE PARTS")) {
									SUPERLARGE_PARTS++;
								}
								if (type.equals("LARGE PARTS")) {
									LARGE_PARTS++;
								}
								if (type.equals("MID PARTS")) {
									MID_PARTS++;
								}
								if (type.equals("SMALL PARTS")) {
									SMALL_PARTS++;
								}
							}
						}

						String[] str = new String[10];// ��10������
						str[0] = Integer.toString(rownum + 1);// ���
						str[1] = PARTS_NAME;

						// SUPERLARGE_PARTS = 0;
						if (SUPERLARGE_PARTS != 0) {
							str[2] = Integer.toString(SUPERLARGE_PARTS);
						}
						// LARGE_PARTS = 0;
						if (LARGE_PARTS != 0) {
							str[3] = Integer.toString(LARGE_PARTS);
						}
						// MID_PARTS = 0;
						if (MID_PARTS != 0) {
							str[4] = Integer.toString(MID_PARTS);
						}
						// SMALL_PARTS = 0;
						if (SMALL_PARTS != 0) {
							str[5] = Integer.toString(SMALL_PARTS);
						}
						if (CLAMP != 0) {
							str[6] = Integer.toString(CLAMP);
						}
						// CLAMP = 0;

//						ArrayList QList = Util.searchBOMLine(bl, "OR",
//								new String[] { "bl_item_object_type", "bl_item_object_type" }, "==",
//								new String[] { "ǹ", "BIW Gun" });
//						ArrayList QList = getGunNums(bl);
//						ArrayList HDList = Util.searchBOMLine(bl, "OR",
//								new String[] { "bl_item_object_type", "bl_item_object_type" }, "==",
//								new String[] { "����", "Weld Point" });
						Integer[] hdnums = getWeldPointNums(bl);
//						if (QList != null) {
//							GUN = QList.size();
//							str[7] = Integer.toString(GUN);
//						} else {
//							str[7] = "0";
//						}
						str[7] = getGunNums(bl).toString();
						if (hdnums != null) {
							PSW = hdnums[0];
							RSW = hdnums[1];
							if (RSW != 0) {
								str[8] = hdnums[1].toString();
							}
							if (PSW != 0) {
								str[9] = hdnums[0].toString();
							}
							hsnum = hsnum + PSW + RSW;
							rswnum = rswnum + RSW;
						}
						rownum++;
						list.add(str);
					}

				}

				// װ����
				if (ajm == 2) {
					int LARGE_PARTS = 0;// LARGE PARTS��ֵ
					int SUPERLARGE_PARTS = 0;// SUPERLARGE_PARTS��ֵ
					int MID_PARTS = 0;// MID PARTS��ֵ
					int SMALL_PARTS = 0;// SMALL PARTS��ֵ
					int CLAMP = 0;// CLAMP& UN-CLAMP��ֵ

					PARTS_NAME = "FIX - " + bl.getProperty("bl_rev_object_name");

					// ���ݺ�װ��λ��ѯ���
					ArrayList part = new ArrayList();// �������

					getSolItmPart(bl, part);

					for (int j = 0; j < part.size(); j++) {
						TCComponentItemRevision rev = (TCComponentItemRevision) part.get(j);
						String partno = rev.getProperty("dfl9_part_no");
						if (partno.length() > 5) {
							partno = partno.substring(0, 5);
						}
						if (rule.containsKey(partno)) {
							String type = rule.get(partno);
							if (type.equals("SUPER LARGE PARTS")) {
								SUPERLARGE_PARTS++;
							}
							if (type.equals("LARGE PARTS")) {
								LARGE_PARTS++;
							}
							if (type.equals("MID PARTS")) {
								MID_PARTS++;
							}
							if (type.equals("SMALL PARTS")) {
								SMALL_PARTS++;
							}
						}
					}

					String[] str = new String[10];// ��10������
					str[0] = Integer.toString(rownum + 1);// ���
					str[1] = PARTS_NAME;

					// SUPERLARGE_PARTS = 0;
					if (SUPERLARGE_PARTS != 0) {
						str[2] = Integer.toString(SUPERLARGE_PARTS);
					}
					// LARGE_PARTS = 0;
					if (LARGE_PARTS != 0) {
						str[3] = Integer.toString(LARGE_PARTS);
					}
					// MID_PARTS = 0;
					if (MID_PARTS != 0) {
						str[4] = Integer.toString(MID_PARTS);
					}
					// SMALL_PARTS = 0;
					if (SMALL_PARTS != 0) {
						str[5] = Integer.toString(SMALL_PARTS);
					}
					if (CLAMP != 0) {
						str[6] = Integer.toString(CLAMP);
					}
					rownum++;

					list.add(str);
				}
			}
			// װ����������check��3��
			if (ajm == 2) {
				list.add(rownum);
			} else {
				list.add(rownum + 3);
			}

			return list;
		} catch (Exception ex) {
			ex.printStackTrace();
		}
		return list;
	}

	private Integer getGunNums(TCComponentBOMLine bl) throws TCException {
		// TODO Auto-generated method stub
		Integer gunnum = 0;
		AIFComponentContext[] children = bl.getChildren();
		for (AIFComponentContext chil : children) {
			TCComponentBOMLine disbl = (TCComponentBOMLine) chil.getComponent();
			TCComponentItemRevision rev = disbl.getItemRevision();
			if (rev.isTypeOf("B8_BIWDiscreteOPRevision")) {
				ArrayList gunlist = Util.getChildrenByBOMLine(disbl, "B8_BIWGunRevision");	
				if(gunlist!=null) {
					gunnum = gunnum + gunlist.size();
				}
			}
		}

		return gunnum;
	}

	/*
	 * ��ȡ��λ�µĻ���������˹����������
	 */
	private Integer[] getWeldPointNums(TCComponentBOMLine bl) throws TCException {
		Integer[] nums = new Integer[2];
		int pswnum = 0;
		int rswnum = 0;
		AIFComponentContext[] children = bl.getChildren();
		for (AIFComponentContext chil : children) {
			TCComponentBOMLine disbl = (TCComponentBOMLine) chil.getComponent();
			TCComponentItemRevision rev = disbl.getItemRevision();
			if (rev.isTypeOf("B8_BIWDiscreteOPRevision")) {
				ArrayList weldlist = Util.getChildrenByBOMLine(disbl, "WeldPointRevision");
				int factnum = 0;
				if (weldlist != null) {
					factnum = weldlist.size();
				}
				String objectname = Util.getProperty(rev, "object_name");
				if (objectname.substring(0, 1).equals("R")) {
					rswnum = rswnum + factnum;
				} else {
					pswnum = pswnum + factnum;
				}
			}
		}
		nums[0] = pswnum;
		nums[1] = rswnum;

		return nums;
	}

	// ��METAL���߷ŵ�������
	private ArrayList getOrderList(List CList) {
		// �ź���Ĳ��߼���
		ArrayList orderList = new ArrayList();
		// METAL����
		ArrayList MList = new ArrayList();
		// ��METAL����
		ArrayList UNMList = new ArrayList();

		for (int i = 0; i < CList.size(); i++) {
			TCComponentBOMLine bl = (TCComponentBOMLine) CList.get(i);
			int n = getIsAJ(bl);
			if (n == 2) {
				MList.add(bl);
			} else {
				UNMList.add(bl);
			}
		}
		for (int j = 0; j < UNMList.size(); j++) {
			orderList.add(UNMList.get(j));
		}
		for (int k = 0; k < MList.size(); k++) {
			orderList.add(MList.get(k));
		}

		return orderList;
	}

	private int getIsAJ(TCComponentBOMLine bl) {
		// TODO Auto-generated method stub
		int type = 1;
		try {
			String objectname = bl.getProperty("bl_rev_object_name");

			if ((objectname.contains("04") && objectname.contains("FM"))
					|| (objectname.contains("07") && objectname.contains("BM"))) {
				type = 1; // ��Ҫ���ֻ����˹���
			} else if (objectname.contains("METAL")) {
				type = 2; // METAl����
			} else {
				type = 3;// ����
			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return type;
	}

	// �жϺ�װ��λ�������Ƿ��ж���㺸����
	private boolean getIsDiscretes(TCComponentBOMLine bl) {
		// TODO Auto-generated method stub
		boolean flag = false;
		try {
			AIFComponentContext[] children = bl.getChildren();

			for (AIFComponentContext chil : children) {
				TCComponentItemRevision rev = ((TCComponentBOMLine) chil.getComponent()).getItemRevision();
				if (rev.isTypeOf("B8_BIWDiscreteOPRevision") || rev.isTypeOf("B8_BIWOperationRevision")) {
					String objectname = Util.getProperty(rev, "object_name");
					if (objectname.substring(0, 1).equals("R")) {
						flag = true;
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return flag;
	}

	private boolean getIsRSW(TCComponentBOMLine bl) {
		// TODO Auto-generated method stub
		boolean flag = false;
		try {
			AIFComponentContext[] children = bl.getChildren();
			int num = 0;
			for (AIFComponentContext chil : children) {
				TCComponentBOMLine downbl = (TCComponentBOMLine) chil.getComponent();
				TCComponentItemRevision rev = downbl.getItemRevision();
				if (rev.isTypeOf("B8_BIWDiscreteOPRevision")) {
					if (downbl.getProperty("bl_rev_object_name").substring(0, 1).equals("R")) {
						flag = true;
						break;
					}
				}
			}
			if (num > 1) {
				return true;
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return flag;
	}

	// ͨ����װ��λ��ѯ�ж������DFL9SolItmPartRevision
	private ArrayList getSolItmPart(TCComponentBOMLine bl, ArrayList part) {
		try {
			AIFComponentContext[] children = bl.getChildren();
			int num = 0;
			for (AIFComponentContext chil : children) {
				TCComponentBOMLine ch = (TCComponentBOMLine) chil.getComponent();
				TCComponentItemRevision rev = ch.getItemRevision();
				if (rev.isTypeOf("DFL9SolItmPartRevision")) {
					part.add(rev);
				} else {
					getSolItmPart(ch, part);
				}
			}
			return part;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return part;
	}

	// ��ѯ��С��������ѡ���ȡ��С��������Ϣ
	private HashMap<String, String> getSizeRule() {
		HashMap<String, String> rule = new HashMap<String, String>();
		try {

			File file = null;
			Workbook workbook = null;
			TCSession session = (TCSession) app.getSession();
			TCPreferenceService preferenceService = session.getPreferenceService();

			String str = preferenceService.getPreferenceDescription("DFL9_straight_sheet_size_rule");
			if (str != null) {
				String value = preferenceService.getStringValue("DFL9_straight_sheet_size_rule");
				if (value != null) {
					TCComponentDatasetType datatype = (TCComponentDatasetType) session.getTypeComponent("Dataset");
					TCComponentDataset dataset = datatype.find(value);
					if (dataset != null) {
						String type = dataset.getType();

						TCComponentTcFile[] files;
						try {
							files = dataset.getTcFiles();
							if (files.length > 0) {
								file = files[0].getFmsFile();
							}
						} catch (TCException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}

						if (file != null) {
							FileInputStream inputStream = new FileInputStream(file);
							if (type.equals("MSExcel")) {
								workbook = new HSSFWorkbook(inputStream);
								rule = parseCoverExcel(workbook);
							}
							if (type.equals("MSExcelX")) {
								workbook = new XSSFWorkbook(inputStream);
								rule = parseCoverExcel(workbook);
							}
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

	private static HashMap<String, String> parseCoverExcel(Workbook workbook) {
		// TODO Auto-generated method stub
		HashMap<String, String> rule = new HashMap<String, String>();
		// ����sheet

		Sheet sheet = workbook.getSheetAt(0);
		// У��sheet�Ƿ�Ϸ�
		if (sheet == null) {
			return null;
		}
		// ��ȡ��һ������
		int firstRowNum = sheet.getFirstRowNum();
		Row firstRow = (Row) sheet.getRow(firstRowNum);
		if (null == firstRow) {
			logger.warning("����Excelʧ�ܣ��ڵ�һ��û�ж�ȡ���κ����ݣ�");
		}

		// ����ÿһ�е����ݣ��������ݶ���
		int rowStart = firstRowNum + 1;
		int rowEnd = sheet.getPhysicalNumberOfRows();
		for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
			Row row = (Row) sheet.getRow(rowNum);
			if (null == row) {
				continue;
			}
			String[] resultData = convertRowToCoverData(row);
			if (null == resultData) {
				logger.warning("�� " + row.getRowNum() + "�����ݲ��Ϸ����Ѻ��ԣ�");
				continue;
			}
			if (resultData[0] != null && !resultData[0].isEmpty()) {
				rule.put(resultData[0], resultData[1]);
			}
		}

		return rule;
	}

	private static String[] convertRowToCoverData(Row row) {
		// TODO Auto-generated method stub
		String[] value = new String[2];
		Cell cell;
		// ���ǰ5λ
		cell = row.getCell(8);
		if (cell != null) {
			String partno = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
			value[0] = partno.trim();
		}
		// �������
		cell = row.getCell(10);
		if (cell != null) {
			String parttype = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
			value[1] = parttype.trim();
		}
		return value;
	}

	private static String convertCellValueToString(Cell cell, int type) {
		if (cell == null) {
			return null;
		}
		String returnValue = null;
		switch (type) {
		case Cell.CELL_TYPE_NUMERIC: // ����
			Double doubleValue = cell.getNumericCellValue();
			// ��ʽ����ѧ��������ȡһλ����
			DecimalFormat df = new DecimalFormat("0.00");
			returnValue = df.format(doubleValue);
			break;
		case Cell.CELL_TYPE_STRING: // �ַ���
			cell.setCellType(Cell.CELL_TYPE_STRING);
			returnValue = cell.getStringCellValue();
			break;
		case Cell.CELL_TYPE_BOOLEAN: // ����
			Boolean booleanValue = cell.getBooleanCellValue();
			returnValue = booleanValue.toString();
			break;
		case Cell.CELL_TYPE_BLANK: // ��ֵ
			break;
		case Cell.CELL_TYPE_FORMULA: // ��ʽ
			returnValue = cell.getCellFormula();
			break;
		case Cell.CELL_TYPE_ERROR: // ����
			break;
		default:
			break;
		}
		return returnValue;
	}
}
