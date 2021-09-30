package com.dfl.report.mfcadd;

import java.io.File;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentDatasetType;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentMEOP;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.kernel.tcservices.TcBOMService;

public class DirectMatSummaryReportOperation {
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
	private HashMap<String, String> hmNesProportion;
	private HashMap<String, String> hmNesPrice;
//	private HashMap<String, String> hmTypeProportion;
//	private HashMap<String, String> hmTypePrice;
	private final int COL_TYPE = 0;
	private final int COL_PROPORTION = 2;
	private final int COL_PRICE = 4;
	private HashMap<TCComponentBOMLine, String> hmLineBody;
	private static final String[] weldProps = new String[] {"b8_modelno", "b8_Long", "b8_LongUOM", "b8_Diameter", "b8_Hight"};
	private String sumMoney = "0";
	public DirectMatSummaryReportOperation(TCComponentBOMLine bop, TCComponentBOMLine[] lines, TCComponent folder) {
		bopLine = bop;
		session = bopLine.getSession();
		virtualLines = lines;
		datasetLocation = folder;
		lstBodies = new ArrayList<String>();
		hmNesProportion = new HashMap<String, String>();
		hmNesPrice = new HashMap<String, String>();
//		hmTypeProportion = new HashMap<String, String>();
//		hmTypePrice = new HashMap<String, String>();
		hmBodyData = new HashMap<String,List< String[]>>();
		hmLineBody = new HashMap<TCComponentBOMLine, String>();
		getAndoutReport();
	}
	public void getAndoutReport() {
		try {
			// 界面显示进度并输出执行步骤
			ReportViwePanel viewPanel = new ReportViwePanel("生成报表");
			viewPanel.setVisible(true);
			viewPanel.addInfomation("正在获取模板...\n", 20, 100);
			String prefValue = session.getPreferenceService().getStringValue("DFL9_DirectMate_CountRule");
//			if(prefValue == null || prefValue.length() == 0) {
//				viewPanel.addInfomation("错误：配置直材计算规则Excel文件的首选项DFL9_DirectMate_CountRule 未配置\n", 100,100);
//				return;
//			}
			String countRule = TemplateUtil.getTemplateFile(prefValue);
//			if (countRule == null) {
//				viewPanel.addInfomation("错误：没有找到直材计算规则的Excel文件，请先在TC中添加，名称为：" + prefValue, 100,100);
//				return;
//			}
			LargeExcelFileReadUtil example = new LargeExcelFileReadUtil();
			String[][] infos = example.getExcelDatas(countRule);
			try {
				if(infos != null && infos.length > 1) {
					for(int i = 1; i < infos.length ; i ++) {
						if(infos[i].length > 0 && infos[i][this.COL_TYPE] != null && infos[i][this.COL_TYPE].length() > 0) {
							if(infos[i].length > 4 && infos[i][this.COL_PRICE] != null && infos[i][this.COL_PRICE].length() > 0) {
								this.hmNesPrice.put(infos[i][this.COL_TYPE], infos[i][this.COL_PRICE]);
							}
							if(infos[i].length > 3 && infos[i][this.COL_PROPORTION] != null && infos[i][this.COL_PROPORTION].length() > 0) {
								this.hmNesProportion.put(infos[i][this.COL_TYPE], infos[i][this.COL_PROPORTION]);
							}
						}
//						if(infos[i][this.COL_TYPE] != null && infos[i][this.COL_TYPE].length() > 0) {
//							if(infos[i][this.COL_PRICE] != null && infos[i][this.COL_PRICE].length() > 0) {
//								this.hmTypePrice.put(infos[i][this.COL_TYPE], infos[i][this.COL_PRICE]);
//							}
//							if(infos[i][this.COL_PROPORTION] != null && infos[i][this.COL_PROPORTION].length() > 0) {
//								this.hmTypeProportion.put(infos[i][this.COL_TYPE], infos[i][this.COL_PROPORTION]);
//							}
//						}
					}
				}
			}catch(Exception e) {
				e.printStackTrace();
			}
			
			
			// 查询并导出模板
			String inputStream = TemplateUtil.getTemplateFile("DFL_Template_TQDirectMetaList");
//			if (inputStream == null) {
//				viewPanel.addInfomation("错误：没有找到直材清单（同期）的模板，请先在TC中添加模板(名称为：DFL_Template_TQDirectMetaList)\n", 100,100);
//				return;
//			}
			viewPanel.addInfomation("开始输出报表...\n", 35, 100);
			String familycode = bopLine.getItemRevision().getProperty("project_ids");// 车型
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
						sb.append(facs[i]);
						break;
					}
				}
				fac = sb.toString();
				line = splits[2].substring(fac.length());
			}
			title = vehicle + "_直材清单_";
			SimpleDateFormat sim = new SimpleDateFormat("yyyy.MM.dd HH时");
			curdate = sim.format(new Date());
			title = title + curdate;
			getReportData(this.bopLine);
			if(this.hmBodyData.size() == 0) {
				viewPanel.addInfomation("未在所选目标下找到直材信息...\n", 100, 100);
				try {
					viewPanel.setVisible(false);
					viewPanel.dispose();
				}catch(Exception e) {
					e.printStackTrace();
				}
				MFCUtility.errorMassges("未在所选目标下找到直材信息 ！");
				return;
			}
			viewPanel.addInfomation("开始写数据，请耐心等待...\n", 70, 100);
			String newName = System.getenv("TMP") + File.separator + MFCUtility.fileNameReplace(title, "") + ".xlsx";
			
			POIExcel poi = new POIExcel();
			poi.specifyTemplate(inputStream);
			poi.fillCellValue(0, 1, vehicle);
			poi.fillCellValue(0, 4, fac);
			poi.fillCellValue(0, 10, line);
			System.out.println("rows := " + rows);
			if(rows > 36) {
				poi.insertRow(36, rows - 36);
			}
			int i = 0;
			int j = 0, k = 0;
			int count = this.lstBodies.size();
			int rowIndex = 2;
			int cntLines = 0;
			int bodyIndex = 2;
			for(i = 0 ; i < count; i ++) {
				String body = this.lstBodies.get(i);
				List<String[]> list = this.hmBodyData.get(body);
				if(list == null) {
					System.out.println("error body data : = " + body);
					continue;
				}
				cntLines = list.size();
				System.out.println("body := " + body + " --> cntLines := " + cntLines);
				for(j = 0; j < cntLines; j ++) {
					String[] rowdata = list.get(j);
					for(k = 1; k < 16; k ++) {
						poi.fillCellValue(rowIndex, k, rowdata[k] == null ? "" : rowdata[k]);
					}
					rowIndex ++;
				}
				if((rowIndex - 1) > bodyIndex) {
					poi.addMergedRegion(bodyIndex, 0, rowIndex - 1, 0);
				}
				if(!body.equals("-")) {
					poi.fillCellValue(bodyIndex, 0, body);
				}
				bodyIndex = rowIndex;
			}
			if(rows > 36) {
				poi.fillCellValue(3 + rows , 15, sumMoney);
			}else {
				poi.fillCellValue(39 , 15, sumMoney);
			}
			
			poi.renameSheet(0, vehicle +"车型计算");
			poi.outputExcel(newName);
			File file = new File(inputStream);
			file.delete();
			inputStream = newName;
			viewPanel.addInfomation("创建数据集，请耐心等待...\n", 90, 100);
			TCComponentDatasetType wordType = (TCComponentDatasetType) bopLine.getSession().getTypeComponent("MSExcelX");
			TCComponentDataset dataset = wordType.create(title, "", "MSExcelX");
			dataset.setFiles(new String[]{ inputStream }, new String[]{ "excel" });
			if(datasetLocation instanceof TCComponentFolder) {
				datasetLocation.add("contents", dataset);
			}else if(datasetLocation instanceof TCComponentItemRevision) {
				datasetLocation.add("IMAN_specification", dataset);
			}
			
			file = new File(inputStream);
			file.delete();
			viewPanel.addInfomation("输出报表完成，请在选择保存的文件夹下查看！...\n", 100, 100);
		}catch(Exception e) {
			e.printStackTrace();
			MFCUtility.errorMassges("异常：" + e.getLocalizedMessage());
		}
	}
	private void getReportData(TCComponentBOMLine pline) {
		try {
			List<TCComponent> lstScope = new ArrayList<TCComponent>();
			if(this.virtualLines != null) {
				for(int i = 0; i < this.virtualLines.length; i ++) {
					if(virtualLines[i] == null) {
						System.out.println("virtualLines[i] == null");
					}
					lstScope.add(virtualLines[i]);
				}
			}else {
				if(pline == null) {
					System.out.println("pline is null");
				}
				lstScope.add(pline);
			}
			List<TCComponent> tcclist1 = Util.callStructureSearch(lstScope, "__DFL_Find_Object_by_Name", new String[] { "NAME", "LX"},
					new String[] { "*涂胶*", "B8_BIWArcWeldOP" });
			List<TCComponent> tcclist2 = Util.callStructureSearch(lstScope, "__DFL_Find_Object_by_Name", new String[] { "NAME", "LX"},
					new String[] { "*弧焊*", "B8_BIWArcWeldOP" });
			System.out.println("涂胶工序：" + tcclist1.size());
			System.out.println("弧焊工序：" + tcclist2.size());
			List<TCComponent> tcclist = new ArrayList<TCComponent>();
			for(int i = 0; i < tcclist1.size() ; i ++) {
				System.out.println("涂胶：" + tcclist1.get(i));
				if(!tcclist.contains(tcclist1.get(i))) {
					tcclist.add(tcclist1.get(i));
				}
			}
			for(int i = 0; i < tcclist2.size() ; i ++) {
				System.out.println("弧焊：" + tcclist2.get(i));
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
					System.out.println("涂胶：" + opLine + " 不是工序类型！");
					continue;
				}
				String body = this.getBodyinfo(opLine);
				if(body == null || body.length() == 0) {
					System.out.println("涂胶：" + opLine + " 工序未得到上层产线信息！");
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
						String[] rowdata = new String[16];
						rowdata[0] = body;
						rowdata[1] = weld.getPropertyDisplayableValue(weldProps[0]);
						rowdata[2] = weld.getPropertyDisplayableValue(weldProps[1]);
						rowdata[2] = StringUtil.getStringCutZero(rowdata[2]);
						rowdata[3] = weld.getPropertyDisplayableValue(weldProps[2]);
						rowdata[4] = weld.getPropertyDisplayableValue(weldProps[3]);
						rowdata[4] = StringUtil.getStringCutZero(rowdata[4]);
						rowdata[5] = "";
						if(!StringUtil.isEmpty(rowdata[4])) {
							rowdata[5] = "mm";
						}
						rowdata[6] = weld.getPropertyDisplayableValue(weldProps[4]);
						rowdata[6] = StringUtil.getStringCutZero(rowdata[6]);
						rowdata[7] = "";
						if(!StringUtil.isEmpty(rowdata[6])) {
							rowdata[7] = "mm";
						}
						rowdata[8] = "";//this.proportion == null ? " " : this.proportion;
						if(!StringUtil.isEmpty(rowdata[1]) && this.hmNesProportion.containsKey(rowdata[1])) {
							rowdata[8] = this.hmNesProportion.get(rowdata[1]);
						}
//						else if(!StringUtil.isEmpty(rowdata[1]) && this.hmTypeProportion.containsKey(rowdata[1])) {
//							rowdata[8] = this.hmTypeProportion.get(rowdata[1]);
//						}
						rowdata[9] = "";
						rowdata[10] = "";
						System.out.println("型号rowdata[1] := " + rowdata[1]);
						System.out.println("直径rowdata[4] := " + rowdata[4]);
						System.out.println("比重rowdata[8] := " + rowdata[8]);
						System.out.println("长度rowdata[2] := " + rowdata[2]);
						System.out.println("高度rowdata[6] := " + rowdata[6]);
						if(rowdata[3].equals("m") || rowdata[3].equals("米")) {
							try {
								rowdata[9] = new BigDecimal(3.1415).multiply(new BigDecimal(rowdata[4]))
										.multiply(new BigDecimal(rowdata[4])).multiply(new BigDecimal(rowdata[8]))
										//.multiply(new BigDecimal(rowdata[2]))
										.divide(new BigDecimal(4), 3, BigDecimal.ROUND_HALF_UP).toString();
								//rowdata[9] = StringUtil.getStringCutZero(rowdata[9]);
							}catch(Exception e) {
								e.printStackTrace();
							}
						}else if(rowdata[3].equals("点")) {
							try {
								rowdata[9] = new BigDecimal(3.1415).multiply(new BigDecimal(rowdata[4]))
										.multiply(new BigDecimal(rowdata[4])).multiply(new BigDecimal(rowdata[8]))
										//.multiply(new BigDecimal(rowdata[2]))
										.multiply(new BigDecimal(rowdata[6]))
										.divide(new BigDecimal(4000), 3, BigDecimal.ROUND_HALF_UP).toString();
								//rowdata[9] = StringUtil.getStringCutZero(rowdata[9]);
							}catch(Exception e) {
								e.printStackTrace();
							}
						}
						if(!StringUtil.isEmpty(rowdata[3])) {
							rowdata[10] = "g/" + rowdata[3];
						}
						rowdata[11] = "";
						try {
							rowdata[11] = new BigDecimal(rowdata[2]).multiply(new BigDecimal(rowdata[9])).
									divide(new BigDecimal(1000), 3, BigDecimal.ROUND_HALF_UP).toString();
							//rowdata[11] = StringUtil.getStringCutZero(rowdata[11]);
						}catch(Exception e) {
							e.printStackTrace();
						}
						rowdata[12] = "";
						if(!StringUtil.isEmpty(rowdata[11])) {
							rowdata[12] = "kg";
						}
						rowdata[13] = "";
						if(!StringUtil.isEmpty(rowdata[11])) {
							try {
								rowdata[13] = new BigDecimal(rowdata[2]).multiply(new BigDecimal(rowdata[9])).
										divide(new BigDecimal(1), 3, BigDecimal.ROUND_HALF_UP).toString();
								//rowdata[13] = StringUtil.getStringCutZero(rowdata[13]);
							}catch(Exception e) {
								e.printStackTrace();
							}
						}
						rowdata[14] = "";//price == null ? "" : price;
						if(!StringUtil.isEmpty(rowdata[1]) && this.hmNesPrice.containsKey(rowdata[1])) {
							rowdata[14] = this.hmNesPrice.get(rowdata[1]);
//							try {
//								rowdata[14] = new BigDecimal(rowdata[14]).
//										divide(new BigDecimal(1), 3, BigDecimal.ROUND_HALF_UP).toString();
//								//rowdata[14] = StringUtil.getStringCutZero(rowdata[14]);
//							}catch(Exception e) {
//								e.printStackTrace();
//							}
						}
//						else if(!StringUtil.isEmpty(rowdata[1]) && this.hmTypePrice.containsKey(rowdata[1])) {
//							rowdata[14] = this.hmTypePrice.get(rowdata[1]);
//						}
						System.out.println("单价rowdata[14] := " + rowdata[14]);
						rowdata[15] = "";
						if(!StringUtil.isEmpty(rowdata[11])) {
							try {
								rowdata[15] = new BigDecimal(rowdata[11]).multiply(new BigDecimal(rowdata[14])).
										divide(new BigDecimal(1), 3, BigDecimal.ROUND_HALF_UP).toString();
								//rowdata[15] = StringUtil.getStringCutZero(rowdata[15]);
							}catch(Exception e) {
								e.printStackTrace();
							}
						}
						if(!StringUtil.isEmpty(rowdata[15])) {
							try {
								this.sumMoney = new BigDecimal(sumMoney).add(new BigDecimal(rowdata[15])).toString();
							}catch(Exception e) {
								e.printStackTrace();
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
}
