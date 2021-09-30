package com.dfl.report.mfcadd;

import java.io.File;
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

public class DATUMSummaryReportOperation {
	TCComponentBOMLine bopLine = null;
	TCComponent datasetLocation = null;
	String title = "";
	String curdate = "";
	private SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy.MM.dd");
	int rows = 0;
	private List<String> lstBodies = null;
	private HashMap<String, List<String>> hmBodyLine;
	private HashMap<String, String[]> hmBLQty;
	private HashMap<String, List<String>> hmBLTimes;
	public DATUMSummaryReportOperation(TCComponentBOMLine bop, TCComponent folder) {
		bopLine = bop;
		datasetLocation = folder;
		lstBodies = new ArrayList<String>();
		hmBodyLine = new HashMap<String, List<String>>();
		hmBLQty = new HashMap<String, String[]>();
		hmBLTimes = new HashMap<String, List<String>>();
		getAndoutReport();
	}
	public void getAndoutReport() {
		try {
			// 界面显示进度并输出执行步骤
			ReportViwePanel viewPanel = new ReportViwePanel("生成报表");
			viewPanel.setVisible(true);
			viewPanel.addInfomation("正在获取模板...\n", 20, 100);
			// 查询并导出模板
			String inputStream = TemplateUtil.getTemplateFile("DFL_Template_DatumStatistics");
//			if (inputStream == null) {
//				viewPanel.addInfomation("错误：没有找到DATUM做成统计信息汇总表的模板，请先在TC中添加模板(名称为：DFL_Template_DatumStatistics)\n", 100,100);
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
			title = vehicle + "车型" + phase + "_DATUM做成统计信息汇总表";
			SimpleDateFormat sim = new SimpleDateFormat("yyyy.MM.dd");
			curdate = sim.format(new Date());
			getReportData(this.bopLine);
			viewPanel.addInfomation("开始写数据，请耐心等待...\n", 70, 100);
			POIExcel poi = new POIExcel();
			poi.specifyTemplate(inputStream);
			poi.fillCellValue(0, 0, title);
			poi.fillCellValue(0, 6, "日期：" + curdate);
			System.out.println("rows := " + rows);
			if(rows > 8) {
				poi.appendRow(10, rows - 8);
			}
			int i = 0;
			int j = 0, k =0;
			int count = this.lstBodies.size();
			int rowIndex = 3;
			int cntLines = 0;
			int row = 0;
			int bodyIndex = 3;
			int lineIndex = 3;
			for(i = 0 ; i < count; i ++) {
				String body = this.lstBodies.get(i);
				
				List<String> lines = this.hmBodyLine.get(body);
				System.out.println("body := " + body + " --> lines := " + lines.size());
				cntLines = lines.size();
				for(j = 0; j < cntLines; j ++) {
					String key = body + "@@@" + lines.get(j);
					List<String> lstFinishDate = this.hmBLTimes.get(key);
					row  = lstFinishDate.size();
					System.out.println("key := " + key + " --> lstFinishDate := " + row);
					if(row == 0) {
						row = 1;
						rowIndex ++;
					}else {
						for(k = 0; k < row; k ++	) {
							System.out.println("rowIndex := " + rowIndex);
							poi.fillCellValue(rowIndex, 5, lstFinishDate.get(k));
							rowIndex ++;
						}
					}
					System.out.println("end rowIndex := " + rowIndex);
					if(row > 1) {
						poi.addMergedRegion(lineIndex, 1, rowIndex-1, 1);
						poi.addMergedRegion(lineIndex, 2, rowIndex-1, 2);
						poi.addMergedRegion(lineIndex, 3, rowIndex-1, 3);
						poi.addMergedRegion(lineIndex, 4, rowIndex-1, 4);
					}
					poi.fillCellValue(lineIndex, 1, lines.get(j));
					String[] qty = this.hmBLQty.get(key);
					poi.fillCellValue(lineIndex, 2, qty[0]);
					poi.fillCellValue(lineIndex, 3, qty[1]);
					poi.fillCellValue(lineIndex, 4, qty[2]);
					lineIndex = rowIndex;
				}
				if((rowIndex-1) >bodyIndex) {
					poi.addMergedRegion(bodyIndex, 0, rowIndex-1, 0);
				}
				poi.fillCellValue(bodyIndex, 0, body);
				bodyIndex = rowIndex;
			}
			String newName = System.getenv("TMP") + File.separator + MFCUtility.fileNameReplace(title, "") + ".xlsx";			
			poi.outputExcel(newName);
			poi.close();
			File ftmp = new File(inputStream);
			ftmp.delete();
			inputStream = newName;
			//poi.outputExcel(inputStream);
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
			for(i = 0; i < count; i ++) {//遍历虚层
				TCComponentBOMLine cline = (TCComponentBOMLine)children[i].getComponent();
				if(cline.getItem().getType().equals("B8_BIWMEProcLine")) {//线体区域
					String name = cline.getItemRevision().getProperty("object_name");
					String body = MFCUtility.transLine2Body(name);
					System.out.println("name := " + name + " --> " + body);
					if(StringUtil.isEmpty(body)) {
						continue;
					}
					if(!this.lstBodies.contains(body)) {
						this.lstBodies.add(body);
					}
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
					String key = body + "@@@" + name;
					System.out.println("key := " + key);
					AIFComponentContext[] lineChildren = cline.getChildren();
					cntStation = lineChildren.length;
					int stationNums = 0;
					int finishNums = 0;
					List<String> lstFinishDate = new ArrayList<String>();
					for(j = 0; j < cntStation; j ++) {//遍历实层
						TCComponentBOMLine realLineLine = (TCComponentBOMLine)lineChildren[j].getComponent();
						System.out.println("realLineLine.getItem().getType() := " + realLineLine.getItem().getType());
						if(realLineLine.getItem().getType().equals("B8_BIWMEProcLine")) {
							AIFComponentContext[] realChildren = realLineLine.getChildren();
							cntReal = realChildren.length;
							for(l = 0; l < cntReal; l ++) {//遍历工位工艺
								TCComponentBOMLine stationLine = (TCComponentBOMLine)realChildren[l].getComponent();
								String type = stationLine.getItem().getType();
								System.out.println("type := " + type);
								if(type.equals("B8_BIWMEProcStat")) {
									stationNums ++;
									AIFComponentContext[] stationChildren = stationLine.getChildren();
									cntGLL = stationChildren.length;
									for(k = 0; k < cntGLL; k ++) {
										TCComponentBOMLine gllLine = (TCComponentBOMLine)stationChildren[k].getComponent();
										System.out.println("gllLine.getItem().getType() := " + gllLine.getItem().getType());
										String gllname = gllLine.getItemRevision().getProperty("object_name");
										if(gllname!=null && gllname.length()>4) {
											if(gllLine.getItem().getType().equals("B8_MPContainer") && "Datum".toUpperCase().equals(gllname.substring(0, 5).toUpperCase())) {
												finishNums ++;
												String stationName = stationLine.getItemRevision().getProperty("object_name");
												String date = dateFormat.format(gllLine.getItemRevision().getDateProperty("last_mod_date"));
												lstFinishDate.add(stationName + " " + date);
												break;
											}
										}
										
									}
								}
							}
						}
					}
					System.out.println("stationNums :-= " + stationNums);
					System.out.println("finishNums :-= " + finishNums);
					System.out.println("lstFinishDate.size :-= " + lstFinishDate.size());
					String[] qty = new String[3];
					qty[0] = String.valueOf(stationNums);
					qty[1] = String.valueOf(finishNums);
					qty[2] = String.valueOf(stationNums - finishNums);
					this.hmBLQty.put(key, qty);
					this.hmBLTimes.put(key, lstFinishDate);
					if(lstFinishDate.size() > 0) {
						rows += lstFinishDate.size();
					}else {
						rows ++;
					}
				}
			}
		}catch(Exception e) {
			e.printStackTrace();
		}
	}

}
