package com.dfl.report.mfcadd;

import java.io.File;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.dfl.report.common.TCComponentUtils;
import com.dfl.report.home.OpenHomeDialog;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentDatasetType;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentMEOP;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.kernel.tcservices.TcBOMService;
import com.teamcenter.services.rac.core._2008_06.DataManagement.CreateResponse;

public class DirectMatWeldSummaryReportOperation {
	TCComponentBOMLine bopLine = null;
	TCComponent datasetLocation = null;
	String title = "";
	String curdate = "";
	private SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy.M.dd");
	int rows = 0;
	private List<String> lstBodies = null;
	private HashMap<String, List<String[]>> hmBodyData;
	TCComponentBOMLine[] virtualLines = null;
	TCSession session = null;
	private HashMap<String, String> hmTypeProportion;
	private HashMap<String, String> hmTypePrice;
	private HashMap<String, String> hmTypeNes;
	private HashMap<String, String> hmNesType;
	private HashMap<String, String> hmNesProportion;
	private HashMap<String, String> hmNesPrice;
	private final int COL_TYPE = 0;
	private final int COL_NESTYPE = 1;
	private final int COL_PROPORTION = 2;
	private final int COL_PRICE = 4;
	private HashMap<TCComponentBOMLine, String> hmLineBody;
	String version = "";
	Map<String, String> mapBodyLineName;
	private static final String[] weldProps = new String[] {"b8_modelno", "b8_Long", "b8_LongUOM", "b8_Diameter", "b8_Hight"};
	public DirectMatWeldSummaryReportOperation(TCComponentBOMLine bop, TCComponentBOMLine[] lines, 
			TCComponent folder, String ver) {
		bopLine = bop;
		session = bopLine.getSession();
		virtualLines = lines;
		datasetLocation = folder;
		version = ver;
		lstBodies = new ArrayList<String>();
		hmTypeProportion = new HashMap<String, String>();
		hmTypePrice = new HashMap<String, String>();
		hmTypeNes = new HashMap<String, String>();
		hmNesProportion = new HashMap<String, String>();
		hmNesPrice = new HashMap<String, String>();
		hmNesType = new HashMap<String, String>();
		hmBodyData = new HashMap<String,List< String[]>>();
		hmLineBody = new HashMap<TCComponentBOMLine, String>();
		mapBodyLineName = new HashMap<String, String>();
		getAndoutReport();
	}
	public void getAndoutReport() {
		try {
			// ������ʾ���Ȳ����ִ�в���
			ReportViwePanel viewPanel = new ReportViwePanel("���ɱ���");
			viewPanel.setVisible(true);
			viewPanel.addInfomation("���ڻ�ȡģ��...\n", 20, 100);
			String prefValue = session.getPreferenceService().getStringValue("DFL9_DirectMate_CountRule");
//			if(prefValue == null || prefValue.length() == 0) {
//				viewPanel.addInfomation("��������ֱ�ļ������Excel�ļ�����ѡ��DFL9_DirectMate_CountRule δ����\n", 100,100);
//				return;
//			}
			String countRule = TemplateUtil.getTemplateFile(prefValue);
//			if (countRule == null) {
//				viewPanel.addInfomation("����û���ҵ�ֱ�ļ�������Excel�ļ���������TC����ӣ�����Ϊ��" + prefValue, 100,100);
//				return;
//			}
			LargeExcelFileReadUtil example = new LargeExcelFileReadUtil();
			String[][] infos = example.getExcelDatas(countRule);
			try {
				if(infos != null && infos.length > 1) {
					for(int i = 1; i < infos.length ; i ++) {
						if(infos[i].length >0 &&infos[i][this.COL_TYPE] != null && infos[i][this.COL_TYPE].length() > 0) {
							if(infos[i].length > 4 && infos[i][this.COL_PRICE] != null && infos[i][this.COL_PRICE].length() > 0) {
								this.hmTypePrice.put(infos[i][this.COL_TYPE], infos[i][this.COL_PRICE]);
							}
							if(infos[i].length > 2 && infos[i][this.COL_PROPORTION] != null && infos[i][this.COL_PROPORTION].length() > 0) {
								this.hmTypeProportion.put(infos[i][this.COL_TYPE], infos[i][this.COL_PROPORTION]);
							}
							if(infos[i].length > 1 && infos[i][this.COL_NESTYPE] != null && infos[i][this.COL_NESTYPE].length() > 0) {
								this.hmTypeNes.put(infos[i][this.COL_TYPE], infos[i][this.COL_NESTYPE]);
							}
						}
						
						if(infos[i].length > 1 && infos[i][this.COL_NESTYPE] != null && infos[i][this.COL_NESTYPE].length() > 0) {
							if(infos[i].length > 4 &&infos[i][this.COL_PRICE] != null && infos[i][this.COL_PRICE].length() > 0) {
								this.hmNesPrice.put(infos[i][this.COL_NESTYPE], infos[i][this.COL_PRICE]);
							}
							if(infos[i].length > 2 && infos[i][this.COL_PROPORTION] != null && infos[i][this.COL_PROPORTION].length() > 0) {
								this.hmNesProportion.put(infos[i][this.COL_NESTYPE], infos[i][this.COL_PROPORTION]);
							}
							if(infos[i].length >0 &&infos[i][this.COL_TYPE] != null && infos[i][this.COL_TYPE].length() > 0) {
								this.hmNesType.put(infos[i][this.COL_NESTYPE], infos[i][this.COL_TYPE]);
							}
						}
					}
				}
			}catch(Exception e) {
				e.printStackTrace();
			}
			
			
			// ��ѯ������ģ��
			String inputStream = TemplateUtil.getTemplateFile("DFL_Template_HZDirectMetaList");
//			if (inputStream == null) {
//				viewPanel.addInfomation("����û���ҵ�ֱ���嵥����������ģ�壬������TC�����ģ��(����Ϊ��DFL_Template_HZDirectMetaList)\n", 100,100);
//				return;
//			}
			viewPanel.addInfomation("��ʼ�������...\n", 35, 100);
			String familycode = bopLine.getItemRevision().getProperty("project_ids");// ����
			String vehicle = Util.getDFLProjectIdVehicle(familycode);
			String factory = "";
			String bopName = bopLine.getItemRevision().getProperty("object_name");
			String[] splits = bopName.split("_");
			String fac = "";
			String line = "";
			if(splits.length > 3) {
				factory = splits[2];
				char[] facs = factory.toCharArray();
				StringBuffer sb = new StringBuffer();
				for(int i = 0; i < facs.length; i ++	) {
					if(facs[i] >= 'A' && facs[i] <= 'Z') {
						sb.append(facs[i]);
					}else if(facs[i]>='1' && facs[i] <= '9') {
						switch(facs[i]) {
						case '1':
							sb.append("һ") ;
							break;
						case '2':
							sb.append("��") ;
							break;
						case '3':
							sb.append("��") ;
							break;
						case '4':
							sb.append("��") ;
							break;
						case '5':
							sb.append("��") ;
							break;
						case '6':
							sb.append("��") ;
							break;
						case '7':
							sb.append("��") ;
							break;
						case '8':
							sb.append("��") ;
							break;
						case '9':
							sb.append("��") ;
							break;
						}
						break;
					}
				}
				fac = sb.toString();
				line = splits[2].substring(fac.length());
				if(line.contains("1")) {
					line = "һ";
				}else if(line.contains("2")) {
					line = "��";
				}else if(line.contains("3")) {
					line = "��";
				}else if(line.contains("4")) {
					line = "��";
				}else if(line.contains("5")) {
					line = "��";
				}else if(line.contains("6")) {
					line = "��";
				}else if(line.contains("7")) {
					line = "��";
				}else if(line.contains("8")) {
					line = "��";
				}else if(line.contains("9")) {
					line = "��";
				}
				
			}
			
			
			title = vehicle + "_" + version + "_ֱ���嵥��������_";
			SimpleDateFormat sim = new SimpleDateFormat("yyyy.MM.dd HHʱ");
			SimpleDateFormat sim2 = new SimpleDateFormat("yyyy��MM��");
			SimpleDateFormat sim3 = new SimpleDateFormat("yyyy.M");
			curdate = sim.format(new Date());
			title = title + curdate;
			getReportData(this.bopLine);
			if(this.hmBodyData.size() == 0) {
				viewPanel.addInfomation("δ����ѡĿ�����ҵ�ֱ����Ϣ...\n", 100, 100);
				try {
					viewPanel.setVisible(false);
					viewPanel.dispose();
				}catch(Exception e) {
					e.printStackTrace();
				}
				MFCUtility.errorMassges("δ����ѡĿ�����ҵ�ֱ����Ϣ ��");
				return;
			}
			String newName = System.getenv("TMP") + File.separator + MFCUtility.fileNameReplace(title, "") + ".xlsx";
			
			viewPanel.addInfomation("��ʼд���ݣ������ĵȴ�...\n", 70, 100);
			POIExcel poi = new POIExcel();
			poi.specifyTemplate(inputStream, 0);
			poi.fillCellValue(5, 3, "      �������̣�" + fac + "������װ����");
			poi.fillCellValue(6, 3, "      ��    �ͣ���ͨ");
			poi.fillCellValue(7, 3, "      ��    �Σ�" + version);
			poi.fillCellValue(8, 3, "      �ļ���ţ���ͨ-" + factory + "-ZC");
			poi.fillCellValue(10, 3, "      �������ڣ�" + sim2.format(new Date()));
			String username = session.getUser().getProperty("user_name");
			poi.fillCellValue(12, 1, username);
			poi.outputExcel(newName);
			poi.close();
			File ftmp = new File(inputStream);
			ftmp.delete();
			inputStream = newName;
			int i = 0;
			int j = 0, k = 0;
			int count = this.lstBodies.size();
			List<String[]> list = new ArrayList<String[]>();
			for(i = 0 ; i < count; i ++) {
				List<String[]> orderlist = this.hmBodyData.get(lstBodies.get(i));
				if(list == null) {
					System.out.println("error body data : = " + lstBodies.get(i));
					continue;
				}
				System.out.println(lstBodies.get(i) + " --> " + orderlist.size());
				list.addAll(orderlist);
			}
			rows = list.size();
			System.out.println("rows := " + rows);
			int pages = rows % 18 == 0 ? rows/18 : rows/18 + 1;
			System.out.println("pages := " + pages);
			if(pages > 1) {
				poi = new POIExcel();
				poi.specifyTemplate(inputStream, 1);
				
				String[] sheetNames = new String[pages];
				for(i = 0; i < pages; i ++) {
					sheetNames[i] = "" + (i + 1);
				}
				poi.cloneTemplate(1, sheetNames);
				poi.outputExcel(inputStream);
				poi.close();
			}
			for(i = 0; i < pages; i ++) {
				poi = new POIExcel();
				poi.specifyTemplate(inputStream, i + 1);
				poi.fillCellValue(3, 11, vehicle);
				poi.fillCellValue(2, 7, sim3.format(new Date()));
				poi.fillCellValue(1, 19, fac + "����");
				poi.fillCellValue(2, 2, username);
				poi.fillCellValue(2, 19, "��װ" + line +  "��");
				poi.fillCellValue(23, 21, version);
				poi.fillCellValue(26, 21, String.valueOf(pages));
				poi.fillCellValue(25, 21, String.valueOf(i + 1));
				for(j = 0; j < 18; j ++) {
					int dataindex = i * 18 + j;
					if(dataindex == rows) {
						break;
					}
					String[] rowdata = list.get(dataindex);
					int rowindex = j + 5;
					poi.fillCellValue(rowindex, 2, rowdata[0]);
					poi.fillCellValue(rowindex, 4, rowdata[1]);
					poi.fillCellValue(rowindex, 6, rowdata[2]);
					poi.fillCellValue(rowindex, 8, rowdata[3]);
					poi.fillCellValue(rowindex, 10, rowdata[4]);
					poi.fillCellValue(rowindex, 11, rowdata[5]);
					poi.fillCellValue(rowindex, 19, rowdata[6]);
					poi.fillCellValue(rowindex, 21, rowdata[7]);
				}
				if(i > 0) {
					poi.zoomSheet(1 + i, 26, 21, 50, true);
				}
				poi.outputExcel(inputStream);
				poi.close();
			}
			viewPanel.addInfomation("�������ݼ��������ĵȴ�...\n", 90, 100);
			TCComponentDatasetType wordType = (TCComponentDatasetType) bopLine.getSession().getTypeComponent("MSExcelX");
			TCComponentDataset dataset = wordType.create(title, "", "MSExcelX");
			dataset.setFiles(new String[]{ inputStream }, new String[]{ "excel" });
			this.saveFiles(dataset);
			
			File file = new File(inputStream);
			file.delete();
			viewPanel.addInfomation("���������ɣ�����ѡ�񱣴���ļ����²鿴��...\n", 100, 100);
			
//			OpenHomeDialog dialog = new OpenHomeDialog(AIFUtility.getActiveDesktop().getsh, session.getUser().getHomeFolder(),session);
//			dialog.open();
//			
//			datasetLocation = dialog.folder;
//			System.out.println("�ļ��У�"+dialog.folder);
//			
//			if(dialog.flag) {
//				return ;
//			}
//			
//			if(datasetLocation == null ) {
//				return ;
//			}
		}catch(Exception e) {
			e.printStackTrace();
			MFCUtility.errorMassges("�쳣��" + e.getLocalizedMessage());
		}
	}
	private void getReportData(TCComponentBOMLine pline) {
		try {
			List<TCComponent> lstScope = new ArrayList<TCComponent>();
			if(this.virtualLines != null) {
				for(int i = 0; i < this.virtualLines.length; i ++) {
					lstScope.add(virtualLines[i]);
				}
			}else {
				lstScope.add(pline);
			}
//			List<TCComponent> tcclist = Util.callStructureSearch(lstScope, "__DFL_Find_Object_by_Name", new String[] { "NAME", "LX"},
//					new String[] { "*Ϳ��*", "B8_BIWArcWeldOP" });
//			System.out.println("Ϳ������" + tcclist.size());
			List<TCComponent> tcclist1 = Util.callStructureSearch(lstScope, "__DFL_Find_Object_by_Name", new String[] { "NAME", "LX"},
					new String[] { "*Ϳ��*", "B8_BIWArcWeldOP" });
			List<TCComponent> tcclist2 = Util.callStructureSearch(lstScope, "__DFL_Find_Object_by_Name", new String[] { "NAME", "LX"},
					new String[] { "*����*", "B8_BIWArcWeldOP" });
			System.out.println("Ϳ������" + tcclist1.size());
			System.out.println("��������" + tcclist2.size());
			List<TCComponent> tcclist = new ArrayList<TCComponent>();
			for(int i = 0; i < tcclist1.size() ; i ++) {
				TCComponentBOMLine line = (TCComponentBOMLine)tcclist1.get(i);
				System.out.println("Ϳ����" + line.getItem().getProperty("item_id") + "/" + line.getItemRevision().getProperty("item_revision_id") + " := " + line);
				if(!tcclist.contains(tcclist1.get(i))) {
					tcclist.add(tcclist1.get(i));
				}
			}
			for(int i = 0; i < tcclist2.size() ; i ++) {
				TCComponentBOMLine line = (TCComponentBOMLine)tcclist2.get(i);
				System.out.println("������" + line.getItem().getProperty("item_id") + "/" + line.getItemRevision().getProperty("item_revision_id") + " := " + line);
				if(!tcclist.contains(tcclist2.get(i))) {
					tcclist.add(tcclist2.get(i));
				}
			}
			if(tcclist == null || tcclist.size() == 0) {
				return;
			}
			
			int i = 0, j = 0; 
			int count = tcclist.size();
			int cntWeld = 0;
			for(i = 0; i < count; i ++) {
				TCComponentBOMLine opLine = (TCComponentBOMLine)tcclist.get(i);
				if(!(opLine.getItem() instanceof TCComponentMEOP)) {
					System.out.println("Ϳ����" + opLine + " ���ǹ������ͣ�");
					continue;
				}
				String opName = opLine.getItemRevision().getProperty("object_name");
				String body = this.getBodyinfo(opLine);
				if(body == null || body.length() == 0) {
					System.out.println("Ϳ����" + opLine + " ����δ�õ��ϲ������Ϣ��");
					continue;
				}
				TcBOMService.expandOneLevel(session, new TCComponentBOMLine[] {opLine});
				AIFComponentContext[] children = opLine.getChildren();
				cntWeld = children.length;
				List<TCComponentItemRevision> lstWeld = new ArrayList<TCComponentItemRevision>();
				for(j = 0; j < cntWeld; j ++) {
					TCComponentBOMLine weldline = (TCComponentBOMLine)children[j].getComponent();
					if(weldline.getItem().getType().equals("ArcWeld")) {
						lstWeld.add(weldline.getItemRevision());
					}
				}
				if(lstWeld.size() > 0 ) {
					MFCUtility.loadProperties(session, lstWeld.toArray(new TCComponentItemRevision[0]), weldProps);
					cntWeld = lstWeld.size();
					for(j = 0; j < cntWeld; j ++) {
						TCComponentItemRevision weld = lstWeld.get(j);
						System.out.println(body + " --> weld is := " + weld);
						String[] rowdata = new String[8];
						rowdata[0] = "";
						rowdata[1] = "";
						rowdata[2] = "";
						String proportion = "";//this.hmTypeProportion.get(rowdata[1]);
						if(opName.contains("Ϳ��")) {
							rowdata[0] = "��";
							rowdata[1] = weld.getPropertyDisplayableValue(weldProps[0]);
							if(!StringUtil.isEmpty(rowdata[1]) && this.hmTypeNes.containsKey(rowdata[1])) {
								rowdata[2] = this.hmTypeNes.get(rowdata[1]);
							}
							if(!StringUtil.isEmpty(rowdata[1]) && this.hmTypeProportion.containsKey(rowdata[1])) {
								proportion = this.hmTypeProportion.get(rowdata[1]);
							}
						}else if(opName.contains("����")) {
							rowdata[0] = "��˿";
							rowdata[2] = weld.getPropertyDisplayableValue(weldProps[0]);
							if(!StringUtil.isEmpty(rowdata[2]) && this.hmNesType.containsKey(rowdata[2])) {
								rowdata[1] = this.hmNesType.get(rowdata[2]) + "��˿";
							}	
							if(!StringUtil.isEmpty(rowdata[2]) && this.hmNesProportion.containsKey(rowdata[2])) {
								proportion = this.hmNesProportion.get(rowdata[2]);
							}
						}
						System.out.println("DFL := " + rowdata[1]);
						System.out.println("NES := " + rowdata[2]);
						rowdata[3] = body;
						rowdata[4] = "g";
						rowdata[5] = "";
						String uom = weld.getPropertyDisplayableValue(weldProps[2]);
						
						String b8_Long = weld.getPropertyDisplayableValue(weldProps[1]);
						String b8_Diameter = weld.getPropertyDisplayableValue(weldProps[3]);
						String b8_Hight = weld.getPropertyDisplayableValue(weldProps[4]);
						System.out.println("ֱ��b8_Diameter := " + b8_Diameter);
						System.out.println("����rowdata[8] := " + proportion);
						System.out.println("����b8_Long := " + b8_Long);
						System.out.println("�߶�b8_Hight := " + b8_Hight);
						if(proportion != null && proportion.length() > 0) {
							if(uom.equals("m") || uom.equals("��")) {
								try {
									rowdata[5] = new BigDecimal(3.1415).multiply(new BigDecimal(b8_Diameter))
											.multiply(new BigDecimal(b8_Diameter)).multiply(new BigDecimal(proportion))
											.multiply(new BigDecimal(b8_Long))
											.divide(new BigDecimal(4), 1, BigDecimal.ROUND_HALF_UP).toString();
									rowdata[5] = StringUtil.getStringCutZero(rowdata[5]);
								}catch(Exception e) {
									e.printStackTrace();
								}
							}else if(uom.equals("��")) {
								try {
									rowdata[5] = new BigDecimal(3.1415).multiply(new BigDecimal(b8_Diameter))
											.multiply(new BigDecimal(b8_Diameter)).multiply(new BigDecimal(proportion))
											.multiply(new BigDecimal(b8_Long))
											.multiply(new BigDecimal(b8_Hight))
											.divide(new BigDecimal(4000), 1, BigDecimal.ROUND_HALF_UP).toString();
									rowdata[5] = StringUtil.getStringCutZero(rowdata[5]);
								}catch(Exception e) {
									e.printStackTrace();
								}
							}
						}
						rowdata[6] = "�й�";
						if(!StringUtil.isEmpty(rowdata[1])) {
							if(rowdata[1].startsWith("SD")) {
								rowdata[6] = "����";
							}else if(rowdata[1].startsWith("ST") || rowdata[1].startsWith("#")) {
								rowdata[6] = "ʱ����";
							}
						}
						rowdata[7] = "���߰�";
						if(mapBodyLineName.containsKey(body)) {
							String lineName = mapBodyLineName.get(body);
							if(lineName.contains("�ذ�")) {
								rowdata[7] = "�ذ��";
							}else if(lineName.contains("��Χ")) {
								rowdata[7] = "��Χ��";
							}else if(lineName.contains("����")) {
								rowdata[7] = "���ǰ�";
							}else if(lineName.contains("����")) {
								rowdata[7] = "���߰�";
							}else if(lineName.contains("����")) {
								rowdata[7] = "���հ�";
							}else {
								rowdata[7] = "���߰�";
							}
						}
						
						if(this.hmBodyData.containsKey(body)) {
							List<String[]> list = this.hmBodyData.get(body);
							list.add(rowdata);
							rows ++;
							this.hmBodyData.put(body, list);
						}else {
							List<String[]> list = new ArrayList<String[]>();
							list.add(rowdata);
							rows ++;
							this.hmBodyData.put(body, list);
						}
						if(!this.lstBodies.contains(body)) {
							this.lstBodies.add(body);
						}
					}
				}
			}
		}catch(Exception e) {
			e.printStackTrace();
		}
	}
	private String getBodyinfo(TCComponentBOMLine opLine) {
		String body = "";
		try {
			TCComponentBOMLine statLine = opLine.parent();
			if(this.hmLineBody.containsKey(statLine)) {
				return this.hmLineBody.get(statLine);
			}
			TCComponentBOMLine lineLine = statLine.parent();
			TcBOMService.expandOneLevel(session, new TCComponentBOMLine[] {lineLine});
			AIFComponentContext[] children = lineLine.getChildren();
			int i = 0;
			int count = children.length;
			int cntStat = 0;
			for(i = 0; i < count; i ++) {
				TCComponentBOMLine sline = (TCComponentBOMLine)children[i].getComponent();
				if(sline.getItem().getType().equals("B8_BIWMEProcStat")) {
					cntStat ++;
					if(cntStat > 1) {
						break;
					}
				}
			}
			if(cntStat == 1) {
				body = lineLine.getItemRevision().getProperty("b8_ChineseName");
				if(!StringUtil.isEmpty(body)) {
					mapBodyLineName.put(body, body);
					System.out.println("body  " + body + " --> chinesename := " + body);
				}
				if(body.length() == 0) {
					body = "-";
				}
				this.hmLineBody.put(lineLine, body);
				this.hmLineBody.put(statLine, body);
//				if(!this.lstBodies.contains(body)) {
//					this.lstBodies.add(body);
//				}
			}else if(cntStat > 1) {
				body = lineLine.getItemRevision().getProperty("b8_ChineseName") + "\n" + statLine.getItemRevision().getProperty("object_name");
				if(!StringUtil.isEmpty(lineLine.getItemRevision().getProperty("b8_ChineseName"))) {
					mapBodyLineName.put(body, lineLine.getItemRevision().getProperty("b8_ChineseName"));
					System.out.println("body  " + body + " --> chinesename := " + lineLine.getItemRevision().getProperty("b8_ChineseName"));
				}
				this.hmLineBody.put(statLine, body);
//				if(!this.lstBodies.contains(body)) {
//					this.lstBodies.add(body);
//				}
			}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return body;
	}
	public void saveFiles(TCComponentDataset ds) {
		try {
				int i = 0;
				Map<String, Object> itemMap = new HashMap<String, Object>();
				Map<String, Object> itemRevisionMap = new HashMap<String, Object>();
				Map<String, Object> itemRevMasterFormMap = new HashMap<String, Object>();
				itemMap.put("item_id", ""); //$NON-NLS-1$ //$NON-NLS-2$
				itemMap.put("object_name", title); //$NON-NLS-1$
				itemMap.put("object_desc", ""); //$NON-NLS-1$
				itemMap.put("object_type", "DFL9MEDocument"); //$NON-NLS-1$
				itemRevisionMap.put("object_type", "DFL9MEDocumentRevision"); //$NON-NLS-1$
				itemRevisionMap.put("object_name", title); //$NON-NLS-1$
				itemRevisionMap.put("dfl9_process_type", "H"); //$NON-NLS-1$
				itemRevisionMap.put("dfl9_process_file_type", "ZC"); //$NON-NLS-1$
				//itemRevisionMap.put("dfl9_vehiclePlant", "docNo"); 
				itemRevMasterFormMap.put("object_type", "DFL9MEDocumentRevisionMaster"); //$NON-NLS-1$
				CreateResponse respose = TCComponentUtils.create(itemMap, itemRevisionMap, itemRevMasterFormMap);
				int num = respose.serviceData.sizeOfCreatedObjects();
				TCComponentItemRevision rev = null;
				TCComponentItem tccomponentitem = null;
				if(num > 0){
					for(i=0;i<num;i++){
						TCComponent comp = respose.serviceData.getCreatedObject(i);
						if(comp instanceof TCComponentItemRevision){
							rev = (TCComponentItemRevision) comp;						
						}else if(comp instanceof TCComponentItem) {
							tccomponentitem = (TCComponentItem)comp;
						}
					}
				}
				// ����ĵ������ݼ��Ĺ�ϵ
				rev.add("IMAN_specification", ds);
				// ��Ӻ�װ��λ���ĵ��Ĺ�ϵ
				if(datasetLocation instanceof TCComponentFolder) {
					datasetLocation.add("contents", tccomponentitem);
				}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
