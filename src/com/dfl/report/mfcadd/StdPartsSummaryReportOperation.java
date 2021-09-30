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

public class StdPartsSummaryReportOperation {
	TCComponentBOMLine bopLine = null;
	TCComponent datasetLocation = null;
	String title = "";
	String curdate = "";
	private SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy.MM.dd");
	int rows = 0;
	private List<String> lstBodies = null;
	private HashMap<String, List<String>> hmBodyLine;
	private HashMap<String, List<String>> hmBodyLineStations;
	private HashMap<String, List<String[]>> hmBLSStdParts;
	public StdPartsSummaryReportOperation(TCComponentBOMLine bop, TCComponent folder) {
		bopLine = bop;
		datasetLocation = folder;
		lstBodies = new ArrayList<String>();
		hmBodyLine = new HashMap<String, List<String>>();
		hmBodyLineStations = new HashMap<String, List<String>>();
		hmBLSStdParts = new HashMap<String, List<String[]>>();
		getAndoutReport();
	}
	public void getAndoutReport() {
		try {
			// 界面显示进度并输出执行步骤
			ReportViwePanel viewPanel = new ReportViwePanel("生成报表");
			viewPanel.setVisible(true);
			viewPanel.addInfomation("正在获取模板...\n", 20, 100);
			// 查询并导出模板
			String inputStream = TemplateUtil.getTemplateFile("DFL_Template_StandardStatistics");
//			if (inputStream == null) {
//				viewPanel.addInfomation("错误：没有找到车型标准件信息汇总表的模板，请先在TC中添加模板(名称为：DFL_Template_StandardStatistics)\n", 100,100);
//				return;
//			}
			viewPanel.addInfomation("开始输出报表...\n", 35, 100);
			String familycode = bopLine.getItemRevision().getProperty("project_ids");// 车型
			String vehicle = Util.getDFLProjectIdVehicle(familycode);
			String phase = "";
			String bopName = bopLine.getItemRevision().getProperty("object_name");
			String[] splits = bopName.split("_");
			if(splits.length > 3) {
				phase = splits[3];
			}
			title = vehicle + "车型" + phase + "_标准件信息汇总表";
			SimpleDateFormat sim = new SimpleDateFormat("yyyy.MM.dd");
			curdate = sim.format(new Date());
			getReportData(this.bopLine);
			viewPanel.addInfomation("开始写数据，请耐心等待...\n", 70, 100);
			POIExcel poi = new POIExcel();
			poi.specifyTemplate(inputStream);
			poi.fillCellValue(0, 0, title);
			poi.fillCellValue(0, 7,  curdate);
			System.out.println("rows := " + rows);
			if(rows > 8) {
				poi.appendRow(10, rows - 8);
			}
			int i = 0, j = 0, k = 0, m = 0;
			int cntBody = 0, cntLine = 0, cntStat = 0, row = 0;
			int rowIndex = 3;
			int bodyIndex = 3, lineIndex = 3, statIndex = 3;
			cntBody = this.lstBodies.size();
			for(i = 0; i  < cntBody; i ++) {
				String body = this.lstBodies.get(i);
				List<String> lstLines = this.hmBodyLine.get(body);
				cntLine = lstLines.size();
				for(j = 0; j < cntLine; j ++) {
					String lineName = lstLines.get(j);
					List<String> lstStations = this.hmBodyLineStations.get(body +"@@@" + lineName);
					cntStat = lstStations.size();
					for(k = 0; k < cntStat; k ++) {
						String statName = lstStations.get(k);
						List<String[]> lstParts = this.hmBLSStdParts.get(body +"@@@" + lineName + "@@@" + statName);
						row = lstParts.size();
						for(m = 0; m < row; m ++) {
							String[] rowdata = lstParts.get(m);
							poi.fillCellValue(rowIndex, 4, rowdata[0]);
							poi.fillCellValue(rowIndex, 5, rowdata[1]);
							poi.fillCellValue(rowIndex, 6, rowdata[2]);
							rowIndex ++;
						}
						if(row > 1) {
							poi.addMergedRegion(statIndex, 2, rowIndex-1, 2);
						}
						poi.fillCellValue(statIndex, 2, statName);
						statIndex = rowIndex;
					}
					if((rowIndex - 1) > lineIndex) {
						poi.addMergedRegion(lineIndex, 1, rowIndex-1, 1);
					}
					poi.fillCellValue(lineIndex, 1, lineName);
					lineIndex = rowIndex;
				}
				if((rowIndex - 1) > bodyIndex) {
					poi.addMergedRegion(bodyIndex, 0, rowIndex-1, 0);
				}
				poi.fillCellValue(bodyIndex, 0, body);
				bodyIndex = rowIndex;
			}
			poi.outputExcel(inputStream);
			poi.close();
			viewPanel.addInfomation("创建数据集，请耐心等待...\n", 90, 100);
			TCComponentDatasetType wordType = (TCComponentDatasetType) bopLine.getSession().getTypeComponent("MSExcelX");
			TCComponentDataset dataset = wordType.create(title, "", "MSExcelX");
			dataset.setFiles(new String[]{ inputStream }, new String[]{ "excel" });
			if(datasetLocation instanceof TCComponentFolder) {
				datasetLocation.add("contents", dataset);
			}else if(datasetLocation instanceof TCComponentItemRevision) {
				datasetLocation.add("IMAN_specification", dataset);
			}
			
			File file = new File(inputStream);
			file.delete();
			viewPanel.addInfomation("输出报表完成，请在选择保存的文件夹下查看！...\n", 100, 100);
		}catch(Exception e) {
			e.printStackTrace();
			MFCUtility.errorMassges("异常：" + e.getLocalizedMessage());
		}
	}
	private void getReportData(TCComponentBOMLine pline) {
		try {
			AIFComponentContext[] children = pline.getChildren();
			int i = 0, j = 0, cntStation = 0, k = 0, cntGLL = 0, l = 0, cntReal = 0;
			int count = children.length;
			//System.out.println("count := " + count);
			for(i = 0; i < count; i ++) {//遍历虚层
				TCComponentBOMLine cline = (TCComponentBOMLine)children[i].getComponent();
				//System.out.println("cline.getItem().getType() := " + cline.getItem().getType());
				if(cline.getItem().getType().equals("B8_BIWMEProcLine")) {//线体区域
					String name = cline.getItemRevision().getProperty("object_name");
					String body = MFCUtility.transLine2Body(name);
					//System.out.println("name := " + name + " --> " + body);
					if(StringUtil.isEmpty(body)) {
						continue;
					}
					String key = body + "@@@" + name;
					AIFComponentContext[] lineChildren = cline.getChildren();
					cntStation = lineChildren.length;
					List<String> lstStationName = new ArrayList<String>();
					boolean hasStd = false;
					for(j = 0; j < cntStation; j ++) {//遍历实层
						TCComponentBOMLine realLineLine = (TCComponentBOMLine)lineChildren[j].getComponent();
						if(realLineLine.getItem().getType().equals("B8_BIWMEProcLine")) {
							AIFComponentContext[] realChildren = realLineLine.getChildren();
							cntReal = realChildren.length;
							for(l = 0; l < cntReal; l ++) {//遍历工位工艺
								TCComponentBOMLine stationLine = (TCComponentBOMLine)realChildren[l].getComponent();
								String type = stationLine.getItem().getType();
								if(type.equals("B8_BIWMEProcStat")) {
									String stationName = stationLine.getItemRevision().getProperty("object_name");
									String partKey = key + "@@@" + stationName;
									List<String[]> lstPartInfo = new ArrayList<String[]>();
									AIFComponentContext[] stationChildren = stationLine.getChildren();
									cntGLL = stationChildren.length;
									System.out.println("工位子行：" + cntGLL);
									HashMap<String, String> hmPartQty = new HashMap<String, String>();
									List<String> lstPartKey = new ArrayList<String>();
									for(k = 0; k < cntGLL; k ++) {
										TCComponentBOMLine gllLine = (TCComponentBOMLine)stationChildren[k].getComponent();
										String bl_usage_address = gllLine.getProperty("bl_usage_address");
										System.out.println(gllLine.getItem().getType() + " --> bl_usage_address：" + bl_usage_address);
										if(gllLine.getItem().getType().equals("DFL9SolItmPart") &&bl_usage_address.length() > 0 && bl_usage_address.startsWith("MU")) {
											String partname = gllLine.getItemRevision().getProperty("object_name");
											System.out.println("partname := " + partname);
											if(partname.toLowerCase().contains("nut") || 
													partname.toLowerCase().contains("bolt") || 
													partname.toLowerCase().contains("stud")) {
												String partkey = partname + "@@@" + gllLine.getItemRevision().getProperty("dfl9_part_no");
												String qty = gllLine.getProperty("bl_quantity");
												if(StringUtil.isEmpty(qty)) {
													qty = "1";
												}
												if(hmPartQty.containsKey(partkey)) {
													String befqty = hmPartQty.get(partkey);
													qty = new BigDecimal(befqty).add(new BigDecimal(qty)).toString();
													hmPartQty.put(partkey, qty);
												}else {
													hmPartQty.put(partkey, qty);
													lstPartKey.add(partkey);
												}
											}
										}else {
											this.getStdPart(gllLine, lstPartKey, hmPartQty);
										}
									}
									cntGLL = lstPartKey.size();
									if(cntGLL > 0) {
										rows += cntGLL;
										hasStd = true;
										lstStationName.add(stationName);
									}
									System.out.println(stationName + " --> hasStd := " + hasStd);
									for(k = 0; k < cntGLL; k ++) {
										String partkey = lstPartKey.get(k);
										String qty = hmPartQty.get(partkey);
										String[] rowdata = new String[3];
										rowdata[0] = partkey.split("@@@")[0];
										rowdata[1] = partkey.split("@@@").length == 2 ? partkey.split("@@@")[1] : "";
										rowdata[2] = qty;
										lstPartInfo.add(rowdata);
									}
									if(cntGLL > 0) {
										this.hmBLSStdParts.put(partKey, lstPartInfo);
									}
								}
							}
						}
					}
					if(hasStd) {
						this.hmBodyLineStations.put(key, lstStationName);
						if(this.hmBodyLine.containsKey(body)) {
							List<String> lst = this.hmBodyLine.get(body);
							if(!lst.contains(name)) {
								lst.add(name);
								this.hmBodyLine.put(body, lst);
							}
						}else {
							List<String> list = new ArrayList<String>();
							list.add(name);
							this.hmBodyLine.put(body, list);
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
	private void getStdPart(TCComponentBOMLine pline, List<String> lstPartKey, HashMap<String, String> hmPartQty) {
		try {
			AIFComponentContext[] children = pline.getChildren();
			int i = 0; 
			int count = children.length;
			for(i = 0; i < count; i ++) {
				TCComponentBOMLine cline = (TCComponentBOMLine)children[i].getComponent();
				String bl_usage_address = cline.getProperty("bl_usage_address");
				if(cline.getItem().getType().equals("DFL9SolItmPart") &&bl_usage_address.length() > 0 && bl_usage_address.startsWith("MU")) {
					String partname = cline.getItemRevision().getProperty("object_name");
					if(partname.toLowerCase().contains("nut") || 
							partname.toLowerCase().contains("bolt") || 
							partname.toLowerCase().contains("stud")) {
						String partkey = partname + "@@@" + cline.getItemRevision().getProperty("dfl9_part_no");
						String qty = cline.getProperty("bl_quantity");
						if(StringUtil.isEmpty(qty)) {
							qty = "1";
						}
						if(hmPartQty.containsKey(partkey)) {
							String befqty = hmPartQty.get(partkey);
							qty = new BigDecimal(befqty).add(new BigDecimal(qty)).toString();
							hmPartQty.put(partkey, qty);
							System.out.println("repeat partkey := " + partkey + " --> qty := " + qty);
						}else {
							hmPartQty.put(partkey, qty);
							lstPartKey.add(partkey);
							System.out.println("partkey := " + partkey + " --> qty := " + qty);
						}
					}
				}else {
					getStdPart(cline, lstPartKey, hmPartQty);
				}
			}
		}catch(Exception e) {
			e.printStackTrace();
		}
	}
}
