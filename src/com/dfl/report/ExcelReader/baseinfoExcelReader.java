package com.dfl.report.ExcelReader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentDatasetType;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentTcFile;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.schemas.internal.core._2011_06.ict.Array;

/*
 * 基本信息读取
 */
public class baseinfoExcelReader {
	private static Logger logger = Logger.getLogger(baseinfoExcelReader.class.getName()); // 日志打印类

	private static final String XLS = "xls";
	private static final String XLSX = "xlsx";
	// 格式化科学计数法，取一位整数
	DecimalFormat publicdf = new DecimalFormat("0");
    DecimalFormat doubledf = new DecimalFormat("0.0");

	/*
	 * ****************************** 获取用于计算焊点一些属性值的参数
	 */
	public static Object[] getCalculationParameter(AbstractAIFUIApplication app,String prefrencename) {
		Object[] obj = new Object[5];
		try {

			File file = null;
			Workbook workbook = null;
			TCSession session = (TCSession) app.getSession();
			TCPreferenceService preferenceService = session.getPreferenceService();

			String str = preferenceService.getPreferenceDescription(prefrencename);
			if (str != null) {
				String value = preferenceService.getStringValue(prefrencename);
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
								obj = parseCoverExcelForpara(workbook);
							}
							if (type.equals("MSExcelX")) {
								workbook = new XSSFWorkbook(inputStream);
								obj = parseCoverExcelForpara(workbook);
							}
						}
					}
				}
			}
			return obj;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return obj;
	}

	/*
	 * **************************** 读取Excel文件计算参数信息
	 */
	public static Object[] parseCoverExcelForpara(Workbook workbook) {
		// TODO Auto-generated method stub
		Object[] obj = new Object[5];
		List<SequenceWeldingConditionList> swc = new ArrayList<SequenceWeldingConditionList>();
		List<SequenceComparisonTable> sct = new ArrayList<SequenceComparisonTable>();
		List<SFSequenceWeldingConditionList> SFswc = new ArrayList<SFSequenceWeldingConditionList>();
		List<RecommendedPressure> rp = new ArrayList<RecommendedPressure>();
		List<CurrentandVoltage> cv = new ArrayList<CurrentandVoltage>();

		/*********************************
		 * 24序列焊接条件设定表 序列号
		 */

		Sheet sheet = workbook.getSheetAt(0);
		// 校验sheet是否合法
		if (sheet == null) {
			return null;
		}
		// 获取第一行数据
		int firstRowNum = sheet.getFirstRowNum();
		Row firstRow = (Row) sheet.getRow(firstRowNum);
		if (null == firstRow) {
			logger.warning("解析Excel失败，在第一行没有读取到任何数据！");
		}

		// 焊接条件数据从11行到36行
		int rowStart = 10;
		int rowEnd = 36;
		for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
			Row row = (Row) sheet.getRow(rowNum);
			if (null == row) {
				continue;
			}
			SequenceWeldingConditionList resultData = convertRowToSWCData(row);
			if (null == resultData) {
				logger.warning("第 " + row.getRowNum() + "行数据不合法，已忽略！");
				continue;
			}
			swc.add(resultData);
		}

		// 电流电压等参数 从41到数据的最后一行
		rowStart = 40;
		rowEnd = sheet.getPhysicalNumberOfRows();
		for (int rowNum = rowStart; rowNum < rowEnd+2; rowNum++) {
			Row row = (Row) sheet.getRow(rowNum);
			if (null == row) {
				continue;
			}
			CurrentandVoltage resultData = convertRowToCVData(row);
			if (null == resultData) {
				logger.warning("第 " + row.getRowNum() + "行数据不合法，已忽略！");
				continue;
			}
			cv.add(resultData);
		}

		/*******************************************/

		/*********************************
		 * 255序列焊接条件设定表
		 */
		Sheet sheet1 = workbook.getSheetAt(1);
		// 校验sheet是否合法
		if (sheet1 == null) {
			return null;
		}
		// 获取第一行数据
		firstRowNum = sheet1.getFirstRowNum();
		firstRow = (Row) sheet1.getRow(firstRowNum);
		if (null == firstRow) {
			logger.warning("解析Excel失败，在第一行没有读取到任何数据！");
		}
		// 数据行从第5行到23行
		rowStart = 4;
		rowEnd = 23;
		for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
			Row row = (Row) sheet1.getRow(rowNum);
			if (null == row) {
				continue;
			}
			SFSequenceWeldingConditionList resultData = convertRowToSFSWCData(row);
			if (null == resultData) {
				logger.warning("第 " + row.getRowNum() + "行数据不合法，已忽略！");
				continue;
			}
			SFswc.add(resultData);
		}

		/*******************************************/

		/*********************************
		 * 推荐加压力
		 */
		Sheet sheet2 = workbook.getSheetAt(2);
		// 校验sheet是否合法
		if (sheet2 == null) {
			return null;
		}
		// 获取第一行数据
		firstRowNum = sheet2.getFirstRowNum();
		firstRow = (Row) sheet2.getRow(firstRowNum);
		if (null == firstRow) {
			logger.warning("解析Excel失败，在第一行没有读取到任何数据！");
		}
		// 数据行从第3行到数据的最后一行
		rowStart = 2;
		rowEnd = sheet2.getPhysicalNumberOfRows();
		for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
			Row row = (Row) sheet2.getRow(rowNum);
			if (null == row) {
				continue;
			}
			RecommendedPressure resultData = convertRowToRPData(row);
			if (null == resultData) {
				logger.warning("第 " + row.getRowNum() + "行数据不合法，已忽略！");
				continue;
			}
			rp.add(resultData);
		}

		/*******************************************/

		/*********************************
		 * 序列对照表
		 */
		Sheet sheet3 = workbook.getSheetAt(3);
		// 校验sheet是否合法
		if (sheet3 == null) {
			return null;
		}
		// 获取第一行数据
		firstRowNum = sheet3.getFirstRowNum();
		firstRow = (Row) sheet3.getRow(firstRowNum);
		if (null == firstRow) {
			logger.warning("解析Excel失败，在第一行没有读取到任何数据！");
		}
		// 数据行从第2行到数据的最后一行
		rowStart = 1;
		rowEnd = sheet3.getPhysicalNumberOfRows();
		//两行两行遍历
		for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
			Row row = (Row) sheet3.getRow(rowNum);
			Row row1 = (Row) sheet3.getRow(rowNum+1);
			if (null == row || row1 == null) {
				continue;
			}
			SequenceComparisonTable resultData = convertRowToSCTData(row,row1);
			if (null == resultData) {
				logger.warning("第 " + row.getRowNum() + "行数据不合法，已忽略！");
				continue;
			}
			sct.add(resultData);
			rowNum++;
		}
		/*******************************************/

		if (swc.size() > 0) {
			obj[0] = swc;
		}
		if (cv.size() > 0) {
			obj[1] = cv;
		}
		if (SFswc.size() > 0) {
			obj[2] = SFswc;
		}
		if (rp.size() > 0) {
			obj[3] = rp;
		}
		if (sct.size() > 0) {
			obj[4] = sct;
		}

		return obj;
	}

	private static SequenceComparisonTable convertRowToSCTData(Row row, Row row1) {
		// TODO Auto-generated method stub
		SequenceComparisonTable sct = new SequenceComparisonTable();
		Map<String,String> map = new HashMap<String,String>();
		Cell cell1; //日产序列
		Cell cell2; //组别内序列
		cell1 = row.getCell(0);
		String parameterGroup = convertCellValueToString(cell1, Cell.CELL_TYPE_STRING);
		if(parameterGroup == null) {
			return null;
		}
		parameterGroup = parameterGroup.replace("group", "");
		sct.setParameterGroup(parameterGroup);
		cell1= row.getCell(2);
		cell2= row1.getCell(2);
		String cvalue1 = convertCellValueToString(cell1, Cell.CELL_TYPE_STRING);
		String cvalue2 = convertCellValueToString(cell2, Cell.CELL_TYPE_STRING);	
		map.put(cvalue1, cvalue2);
		
		cell1= row.getCell(3);
		cell2= row1.getCell(3);
		String dvalue1 = convertCellValueToString(cell1, Cell.CELL_TYPE_STRING);
		String dvalue2 = convertCellValueToString(cell2, Cell.CELL_TYPE_STRING);
		map.put(dvalue1, dvalue2);
		
		cell1= row.getCell(4);
		cell2= row1.getCell(4);
		String evalue1 = convertCellValueToString(cell1, Cell.CELL_TYPE_STRING);
		String evalue2 = convertCellValueToString(cell2, Cell.CELL_TYPE_STRING);
		map.put(evalue1, evalue2);
		
		cell1= row.getCell(5);
		cell2= row1.getCell(5);
		String fvalue1 = convertCellValueToString(cell1, Cell.CELL_TYPE_STRING);
		String fvalue2 = convertCellValueToString(cell2, Cell.CELL_TYPE_STRING);
		map.put(fvalue1, fvalue2);
		
		cell1= row.getCell(6);
		cell2= row1.getCell(6);
		String gvalue1 = convertCellValueToString(cell1, Cell.CELL_TYPE_STRING);
		String gvalue2 = convertCellValueToString(cell2, Cell.CELL_TYPE_STRING);
		map.put(gvalue1, gvalue2);
		
		cell1= row.getCell(7);
		cell2= row1.getCell(7);
		String hvalue1 = convertCellValueToString(cell1, Cell.CELL_TYPE_STRING);
		String hvalue2 = convertCellValueToString(cell2, Cell.CELL_TYPE_STRING);
		map.put(hvalue1, hvalue2);
		
		cell1= row.getCell(8);
		cell2= row1.getCell(8);
		String ivalue1 = convertCellValueToString(cell1, Cell.CELL_TYPE_STRING);
		String ivalue2 = convertCellValueToString(cell2, Cell.CELL_TYPE_STRING);
		map.put(ivalue1, ivalue2);
		
		cell1= row.getCell(9);
		cell2= row1.getCell(9);
		String jvalue1 = convertCellValueToString(cell1, Cell.CELL_TYPE_STRING);
		String jvalue2 = convertCellValueToString(cell2, Cell.CELL_TYPE_STRING);
		map.put(jvalue1, jvalue2);
		
		cell1= row.getCell(10);
		cell2= row1.getCell(10);
		String kvalue1 = convertCellValueToString(cell1, Cell.CELL_TYPE_STRING);
		String kvalue2 = convertCellValueToString(cell2, Cell.CELL_TYPE_STRING);
		map.put(kvalue1, kvalue2);
		
		cell1= row.getCell(11);
		cell2= row1.getCell(11);
		String lvalue1 = convertCellValueToString(cell1, Cell.CELL_TYPE_STRING);
		String lvalue2 = convertCellValueToString(cell2, Cell.CELL_TYPE_STRING);
		map.put(lvalue1, lvalue2);
		
		cell1= row.getCell(12);
		cell2= row1.getCell(12);
		String mvalue1 = convertCellValueToString(cell1, Cell.CELL_TYPE_STRING);
		String mvalue2 = convertCellValueToString(cell2, Cell.CELL_TYPE_STRING);
		map.put(mvalue1, mvalue2);
		
		cell1= row.getCell(13);
		cell2= row1.getCell(13);
		String nvalue1 = convertCellValueToString(cell1, Cell.CELL_TYPE_STRING);
		String nvalue2 = convertCellValueToString(cell2, Cell.CELL_TYPE_STRING);
		map.put(nvalue1, nvalue2);
		
		cell1= row.getCell(14);
		cell2= row1.getCell(14);
		String ovalue1 = convertCellValueToString(cell1, Cell.CELL_TYPE_STRING);
		String ovalue2 = convertCellValueToString(cell2, Cell.CELL_TYPE_STRING);
		map.put(ovalue1, ovalue2);
		
		cell1= row.getCell(15);
		cell2= row1.getCell(15);
		String pvalue1 = convertCellValueToString(cell1, Cell.CELL_TYPE_STRING);
		String pvalue2 = convertCellValueToString(cell2, Cell.CELL_TYPE_STRING);
		map.put(pvalue1, pvalue2);
		
		cell1= row.getCell(16);
		cell2= row1.getCell(16);
		String qvalue1 = convertCellValueToString(cell1, Cell.CELL_TYPE_STRING);
		String qvalue2 = convertCellValueToString(cell2, Cell.CELL_TYPE_STRING);
		map.put(qvalue1, qvalue2);
		
		System.out.println("序列对照表:" + map );
		
		sct.setValues(map);
			
		return sct;
	}

	private static RecommendedPressure convertRowToRPData(Row row) {
		// TODO Auto-generated method stub
		RecommendedPressure swc = new RecommendedPressure();
		Cell cell;
		cell = row.getCell(0);
		String basethickness = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		if (basethickness == null) {
			return null;
		}
		swc.setBasethickness(basethickness.replace(" ",""));
		cell = row.getCell(1);
		String bvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setBvalue(bvalue);
		cell = row.getCell(2);
		String cvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setCvalue(cvalue);
		cell = row.getCell(3);
		String dvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setDvalue(dvalue);
		cell = row.getCell(4);
		String evalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setEvalue(evalue);
		cell = row.getCell(5);
		String fvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setFvalue(fvalue);
		
		
		System.out.println("推荐加压力:" + basethickness + " " + bvalue + " " + cvalue + " " + dvalue + " " + evalue + " " + fvalue );
		
		return swc;
	}

	private static SFSequenceWeldingConditionList convertRowToSFSWCData(Row row) {
		// TODO Auto-generated method stub
		SFSequenceWeldingConditionList swc = new SFSequenceWeldingConditionList();
		Cell cell;
		cell = row.getCell(0);
		String basethickness = convertCellValueToString(cell, Cell.CELL_TYPE_NUMERIC);
		if (basethickness == null) {
			return null;
		}
		swc.setBasethickness(basethickness);
		cell = row.getCell(1);
		String bvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setBvalue(bvalue);
		cell = row.getCell(2);
		String cvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setCvalue(cvalue);
		cell = row.getCell(3);
		String dvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setDvalue(dvalue);
		cell = row.getCell(4);
		String evalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setEvalue(evalue);
		cell = row.getCell(5);
		String fvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setFvalue(fvalue);
		cell = row.getCell(6);
		String gvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setGvalue(gvalue);
		cell = row.getCell(7);
		String hvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setHvalue(hvalue);
		cell = row.getCell(8);
		String ivalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setIvalue(ivalue);
		cell = row.getCell(9);
		String jvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setJvalue(jvalue);
		cell = row.getCell(10);
		String kvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setKvalue(kvalue);
		cell = row.getCell(11);
		String lvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setLvalue(lvalue);
		cell = row.getCell(12);
		String mvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setMvalue(mvalue);
		
		System.out.println("255序列焊接条件设定表:" + basethickness + " " + bvalue + " " + cvalue + " " + dvalue + " " + evalue + " " + fvalue
				+ " " + gvalue + " " + hvalue + " " + ivalue + " " + jvalue + " " + kvalue + " " + lvalue + " " + mvalue );

		return swc;
	}

	private static CurrentandVoltage convertRowToCVData(Row row) {
		// TODO Auto-generated method stub
		CurrentandVoltage swc = new CurrentandVoltage();
		Cell cell;
		cell = row.getCell(0);
		String sequenceNo = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		if (sequenceNo == null) {
			return null;
		}
		swc.setSequenceNo(sequenceNo);
		cell = row.getCell(1);
		String bvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setBvalue(bvalue);
		cell = row.getCell(2);
		String cvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setCvalue(cvalue);
		cell = row.getCell(3);
		String dvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setDvalue(dvalue);
		cell = row.getCell(4);
		String evalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setEvalue(evalue);
		cell = row.getCell(5);
		String fvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setFvalue(fvalue);
		cell = row.getCell(6);
		String gvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setGvalue(gvalue);
		cell = row.getCell(7);
		String hvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setHvalue(hvalue);
		cell = row.getCell(8);
		String ivalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setIvalue(ivalue);
		cell = row.getCell(9);
		String jvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setJvalue(jvalue);
		cell = row.getCell(10);
		String kvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setKvalue(kvalue);
		cell = row.getCell(11);
		String lvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setLvalue(lvalue);
		cell = row.getCell(12);
		String mvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setMvalue(mvalue);
		cell = row.getCell(13);
		String nvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setNvalue(nvalue);
		
		System.out.println("24序列焊接条件设定表2:" + sequenceNo + " " + bvalue + " " + cvalue + " " + dvalue + " " + evalue + " " + fvalue
				+ " " + gvalue + " " + hvalue + " " + ivalue + " " + jvalue + " " + kvalue + " " + lvalue + " " + mvalue + " " + nvalue);
		

		return swc;
	}

	private static SequenceWeldingConditionList convertRowToSWCData(Row row) {
		// TODO Auto-generated method stub
		SequenceWeldingConditionList swc = new SequenceWeldingConditionList();
		Cell cell;
		cell = row.getCell(0);
		String basethickness = convertCellValueToString(cell, Cell.CELL_TYPE_NUMERIC);
		if (basethickness == null) {
			return null;
		}
		swc.setBasethickness(basethickness);
		cell = row.getCell(1);
		String bvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setBvalue(bvalue);
		cell = row.getCell(2);
		String cvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setCvalue(cvalue);
		cell = row.getCell(3);
		String dvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setDvalue(dvalue);
		cell = row.getCell(4);
		String evalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setEvalue(evalue);
		cell = row.getCell(5);
		String fvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setFvalue(fvalue);
		cell = row.getCell(6);
		String gvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setGvalue(gvalue);
		cell = row.getCell(7);
		String hvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setHvalue(hvalue);
		cell = row.getCell(8);
		String ivalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setIvalue(ivalue);
		cell = row.getCell(9);
		String jvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setJvalue(jvalue);
		cell = row.getCell(10);
		String kvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setKvalue(kvalue);
		cell = row.getCell(11);
		String lvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setLvalue(lvalue);
		cell = row.getCell(12);
		String mvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setMvalue(mvalue);
		cell = row.getCell(13);
		String nvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setNvalue(nvalue);
		cell = row.getCell(14);
		String ovalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setOvalue(ovalue);
		cell = row.getCell(15);
		String pvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setPvalue(pvalue);
		cell = row.getCell(16);
		String qvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setQvalue(qvalue);
		cell = row.getCell(17);
		String rvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setRvalue(rvalue);
		cell = row.getCell(18);
		String svalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setSvalue(svalue);
		cell = row.getCell(19);
		String tvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setTvalue(tvalue);
		cell = row.getCell(20);
		String uvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setUvalue(uvalue);
		cell = row.getCell(21);
		String vvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setVvalue(vvalue);
		cell = row.getCell(22);
		String wvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setWvalue(wvalue);
		cell = row.getCell(23);
		String xvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setXvalue(xvalue);
		cell = row.getCell(24);
		String yvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setYvalue(yvalue);
		cell = row.getCell(25);
		String zvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setZvalue(zvalue);
		cell = row.getCell(26);
		String aAvalue = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		swc.setAAvalue(aAvalue);

		System.out.println("24序列焊接条件设定表:" + basethickness + " " + bvalue + " " + cvalue + " " + dvalue + " " + evalue + " " + fvalue
				+ " " + gvalue + " " + hvalue + " " + ivalue + " " + jvalue + " " + kvalue + " " + lvalue + " " + mvalue + " " + nvalue
				+ " " + ovalue + " " + pvalue + " " + qvalue + " " + rvalue + " " + svalue + " " + tvalue + " " + uvalue + " " + vvalue 
				+ " " + wvalue + " " + xvalue+ " " + yvalue+ " " + zvalue+ " " + aAvalue);
		
		return swc;
	}

	/**
	 * 根据文件后缀名类型获取对应的工作簿对象
	 * 
	 * @param inputStream 读取文件的输入流
	 * @param fileType    文件后缀名类型（xls或xlsx）
	 * @return 包含文件数据的工作簿对象
	 * @throws IOException
	 */
	public static Workbook getWorkbook(InputStream inputStream, String fileType) throws IOException {
		Workbook workbook = null;
		if (fileType.equalsIgnoreCase(XLS)) {
			workbook = new HSSFWorkbook(inputStream);
		} else if (fileType.equalsIgnoreCase(XLSX)) {
			workbook = new XSSFWorkbook(inputStream);
		}
		return workbook;
	}

	/*
	 * 读取Excel文件内容 获取文件流
	 * 
	 * @param tcc 读取的对象
	 * 
	 * @param RelateType 对象下关系
	 * 
	 * @param dataname 数据集的名称
	 */
	public static InputStream getFileinbyreadExcel(TCComponent tccomponent, String RelateType, String dataname) {
		InputStream filein = null;
		File file = null;
		TCComponentDataset basicdata = null;
		TCComponent[] tccs;
		try {
			if (tccomponent != null) {
				tccs = tccomponent.getRelatedComponents(RelateType);
				for (TCComponent item : tccs) {
					if (item instanceof TCComponentDataset) {
						String objectname = item.getProperty("object_name");
						if (Util.formatString(objectname).contains(Util.formatString(dataname)) || Util.formatString(objectname).equals(Util.formatString(dataname))) {
							basicdata = (TCComponentDataset) item;
							break;
						}
					}
				}
			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		if (basicdata != null) {
			String type = basicdata.getType();
			if (type.equals("MSExcelX")) {
				TCComponentTcFile[] files;
				try {
					files = basicdata.getTcFiles();
					if (files.length > 0) {
						file = files[0].getFmsFile();
					}
				} catch (TCException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
		if (file != null) {
			try {
				filein = new FileInputStream(file);
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

		return filein;
	}

	public static InputStream getFileinbyreadExcel2(TCComponent tccomponent, String RelateType, String dataname) {
		InputStream filein = null;
		File file = null;
		TCComponentDataset basicdata = null;
		TCComponentItemRevision rev = null;
		TCComponent[] tccs;
		try {
			if (tccomponent != null) {
				tccs = tccomponent.getRelatedComponents(RelateType);
				for (TCComponent item : tccs) {
					if (item instanceof TCComponentItem) {
						String objectname = item.getProperty("object_name");
						String type = item.getType();
						if (type.equals("DFL9MEDocument") && Util.formatString(objectname).equals(Util.formatString(dataname))) {
							TCComponentItem itemtcc = (TCComponentItem) item;
							rev = itemtcc.getLatestItemRevision();
							break;
						}
					}
				}
			}
			if (rev != null) {
				tccs = rev.getRelatedComponents("IMAN_specification");
				for (TCComponent item : tccs) {
					if (item instanceof TCComponentDataset) {
						String objectname = item.getProperty("object_name");
						if (Util.formatString(objectname).equals(Util.formatString(dataname))) {
							basicdata = (TCComponentDataset) item;
							break;
						}
					}
				}
			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		if (basicdata != null) {
			String type = basicdata.getType();
			if (type.equals("MSExcelX")) {
				TCComponentTcFile[] files;
				try {
					files = basicdata.getTcFiles();
					if (files.length > 0) {
						file = files[0].getFmsFile();
					}
				} catch (TCException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
		if (file != null) {
			try {
				filein = new FileInputStream(file);
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

		return filein;
	}

	/**
	 * 读取Excel文件焊点板组信息
	 * 
	 * @param fileName 要读取的Excel文件所在路径
	 * @return 读取结果列表，读取失败时返回null
	 */
	public static List<WeldPointBoardInformation> readHDExcel(InputStream filein, String fileType) {

		Workbook workbook = null;
		InputStream inputStream = null;

		try {
//			// 获取Excel后缀名
//			String fileType = fileName.substring(fileName.lastIndexOf(".") + 1, fileName.length());
//			// 获取Excel文件
//			File excelFile = new File(fileName);
//			if (!excelFile.exists()) {
//				logger.warning("指定的Excel文件不存在！");
//				return null;
//			}
//			// 获取Excel工作簿
//			inputStream = new FileInputStream(excelFile);
			inputStream = filein;

			if (inputStream == null) {
				logger.warning("指定的Excel文件不存在！");
				return null;
			}
			workbook = getWorkbook(inputStream, fileType);

			// 读取excel中的数据
			List<WeldPointBoardInformation> resultDataList = parseExcel(workbook);

			return resultDataList;
		} catch (Exception e) {
			logger.warning("解析Excel失败，文件名： 错误信息：" + e.getMessage());
			e.printStackTrace();
			return null;
		} finally {
			try {
//                if (null != workbook) {
//                    workbook.close();
//                }
				if (null != inputStream) {
					inputStream.close();
				}
			} catch (Exception e) {
				logger.warning("关闭数据流出错！错误信息：" + e.getMessage());
				return null;
			}
		}
	}

	/*
	 * 读取板组信息
	 */
	public static List<BoardInformation> readBZExcel(InputStream filein, String fileType) {

		Workbook workbook = null;
		InputStream inputStream = null;

		try {
//			// 获取Excel后缀名
//			String fileType = fileName.substring(fileName.lastIndexOf(".") + 1, fileName.length());
//			// 获取Excel文件
//			File excelFile = new File(fileName);
//			if (!excelFile.exists()) {
//				logger.warning("指定的Excel文件不存在！");
//				return null;
//			}
//			// 获取Excel工作簿
//			inputStream = new FileInputStream(excelFile);
			inputStream = filein;
			if (inputStream == null) {
				logger.warning("指定的Excel文件不存在！");
				return null;
			}

			workbook = getWorkbook(inputStream, fileType);

			// 读取excel中的数据
			List<BoardInformation> resultDataList = parseBoradExcel(workbook);

			return resultDataList;
		} catch (Exception e) {
			logger.warning("解析Excel失败，错误信息：" + e.getMessage());
			return null;
		} finally {
			try {
//                if (null != workbook) {
//                    workbook.close();
//                }
				if (null != inputStream) {
					inputStream.close();
				}
			} catch (Exception e) {
				logger.warning("关闭数据流出错！错误信息：" + e.getMessage());
				return null;
			}
		}
	}

	/*
	 * 读取焊点清单信息
	 */
	public static List<WeldPointInfo> readWeldExcel(XSSFWorkbook book, String fileType) {

		Workbook workbook = null;

		try {
			workbook = book;

			// 读取excel中的数据
			List<WeldPointInfo> resultDataList = parseWeldExcel(workbook);

			return resultDataList;
		} catch (Exception e) {
			e.printStackTrace();
			logger.warning("解析Excel失败，错误信息：" + e.getMessage());
			return null;
		} finally {
		}
	}

	/*
	 * 读取基本信息中的板组信息，用于判断是否为第一次写入板组信息
	 */
	public static List<BoardInformation> readBaordExcel(XSSFWorkbook book, String fileType) {

		Workbook workbook = null;

		try {
			workbook = book;

			// 读取excel中的数据
			List<BoardInformation> resultDataList = parseBoradExcel(workbook);

			return resultDataList;
		} catch (Exception e) {
			e.printStackTrace();
			logger.warning("解析Excel失败，错误信息：" + e.getMessage());
			return null;
		} finally {
		}
	}

	/*
	 * 读取封面信息
	 */
	public static List<CoverInfomation> readCoverExcel(InputStream filein, String fileType) {

		Workbook workbook = null;
		InputStream inputStream = null;

		try {
			inputStream = filein;
			if (inputStream == null) {
				logger.warning("指定的Excel文件不存在！");
				return null;
			}
			workbook = getWorkbook(inputStream, fileType);
			// 读取excel中的数据
			List<CoverInfomation> resultDataList = parseCoverExcel(workbook);
			return resultDataList;
		} catch (Exception e) {
			e.printStackTrace();
			logger.warning("解析Excel失败，错误信息：" + e.getMessage());
			return null;
		} finally {
		}
	}

	private static List<CoverInfomation> parseCoverExcel(Workbook workbook) {
		// TODO Auto-generated method stub
		List<CoverInfomation> resultDataList = new ArrayList<>();
		// 解析sheet

		Sheet sheet = workbook.getSheetAt(0);
		// 校验sheet是否合法
		if (sheet == null) {
			return null;
		}
		// 获取第一行数据
		int firstRowNum = sheet.getFirstRowNum();
		Row firstRow = (Row) sheet.getRow(firstRowNum);
		if (null == firstRow) {
			logger.warning("解析Excel失败，在第一行没有读取到任何数据！");
		}

		// 解析每一行的数据，构造数据对象
		int rowStart = 7;
		int rowEnd = 8;
		Row row = (Row) sheet.getRow(rowStart);
		Row row1 = (Row) sheet.getRow(rowEnd);
		CoverInfomation resultData = convertRowToCoverData(row,row1);
		if (null == resultData) {
			logger.warning("第 " + row.getRowNum() + "行数据不合法，已忽略！");
		}
		resultDataList.add(resultData);

		return resultDataList;
	}

	private static CoverInfomation convertRowToCoverData(Row row,Row row1) {
		// TODO Auto-generated method stub
		CoverInfomation resultData = new CoverInfomation();
		Cell cell;
		// 工厂线体
		cell = row1.getCell(2);
		String factoryline = convertCellValueToString(cell);
		if (factoryline == null) {
			factoryline = "";
		}
		String[] str = factoryline.split("-");
		String[] str2 = factoryline.split("：");
		if (str.length > 2) {
			if (str[1].length() > 3) {
				String factory = str[1].substring(0, 3);
				String linebody = str[1].substring(str[1].length() - 1);
				resultData.setFactory(factory);
				resultData.setLinebody(linebody);
			}
		}
		if(str2.length>1) {
			resultData.setFilecode(str2[1]);
		}
		//版次
		cell = row.getCell(2);
		String eiditon = convertCellValueToString(cell);
		if(eiditon == null) {
			eiditon = "";
		}
		if(eiditon.length()>0) {
			String[] strVal = eiditon.split("：");
			if(strVal.length>1) {
				eiditon = strVal[1];
			}
		}
		resultData.setEdition(eiditon);
		return resultData;
	}

	/**
	 * 解析Excel数据
	 * 
	 * @param workbook Excel工作簿中的板组信息sheet对象
	 * @return 解析结果
	 */
	private static List<BoardInformation> parseBoradExcel(Workbook workbook) {
		List<BoardInformation> resultDataList = new ArrayList<>();
		// 解析sheet

		Sheet sheet = workbook.getSheetAt(2);
		// 校验sheet是否合法
		if (sheet == null) {
			return null;
		}
		// 获取第一行数据
		int firstRowNum = sheet.getFirstRowNum();
		Row firstRow = (Row) sheet.getRow(firstRowNum);
		if (null == firstRow) {
			logger.warning("解析Excel失败，在第一行没有读取到任何数据！");
		}

		// 解析每一行的数据，构造数据对象
		int rowStart = firstRowNum + 1;
		int rowEnd = sheet.getPhysicalNumberOfRows();
		for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
			Row row = (Row) sheet.getRow(rowNum);
			if (null == row) {
				continue;
			}
			BoardInformation resultData = convertRowToBoradData(row);
			if (null == resultData) {
				logger.warning("第 " + row.getRowNum() + "行数据不合法，已忽略！");
				continue;
			}
			resultDataList.add(resultData);
		}

		return resultDataList;
	}

	private static List<WeldPointInfo> parseWeldExcel(Workbook workbook) {
		List<WeldPointInfo> resultDataList = new ArrayList<>();
		// 解析sheet

		Sheet sheet = workbook.getSheetAt(3);
		// 校验sheet是否合法
		if (sheet == null) {
			return null;
		}
		// 获取第一行数据
		int firstRowNum = sheet.getFirstRowNum();
		Row firstRow = (Row) sheet.getRow(firstRowNum);
		if (null == firstRow) {
			logger.warning("解析Excel失败，在第一行没有读取到任何数据！");
		}

		// 解析每一行的数据，构造数据对象
		int rowStart = firstRowNum + 1;
		int rowEnd = sheet.getPhysicalNumberOfRows();
		for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
			Row row = (Row) sheet.getRow(rowNum);
			if (null == row) {
				continue;
			}
			List<WeldPointInfo> resultData = convertRowToWeldData(row);
			if (null == resultData) {
				logger.warning("第 " + row.getRowNum() + "行数据不合法，已忽略！");
				continue;
			}
			if (resultData != null && resultData.size() > 0) {
				for (int i = 0; i < resultData.size(); i++) {
					WeldPointInfo wp = resultData.get(i);
					resultDataList.add(wp);
				}

			}

		}

		return resultDataList;
	}

	/**
	 * 解析Excel数据
	 * 
	 * @param workbook Excel工作簿中的焊点信息sheet对象
	 * @return 解析结果
	 */
	private static List<WeldPointBoardInformation> parseExcel(Workbook workbook) {
		List<WeldPointBoardInformation> resultDataList = new ArrayList<>();
		// 解析sheet

		// 只需要取焊点信息和板组清单，先取板组信息，再取焊点信息
		List<BoardInformation> boradlist = parseBoradExcel(workbook);

		Sheet sheet = workbook.getSheetAt(3);
		// 校验sheet是否合法
		if (sheet == null) {
			return null;
		}
		// 获取第一行数据
		int firstRowNum = sheet.getFirstRowNum();
		Row firstRow = (Row) sheet.getRow(firstRowNum);
		if (null == firstRow) {
			logger.warning("解析Excel失败，在第一行没有读取到任何数据！");
		}

		// 解析每一行的数据，构造数据对象
		int rowStart = firstRowNum + 1;
		int rowEnd = sheet.getPhysicalNumberOfRows();
		for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
			Row row = (Row) sheet.getRow(rowNum);
			if (null == row) {
				continue;
			}
			WeldPointBoardInformation resultData = convertRowToData(row, boradlist);
			if (null == resultData) {
				logger.warning("第 " + row.getRowNum() + "行数据不合法，已忽略！");
				continue;
			}
			if (resultData != null) {
				resultDataList.add(resultData);
			}
		}
		return resultDataList;
	}

	/**
	 * 将单元格内容转换为字符串
	 * 
	 * @param cell
	 * @return
	 */
	private static String convertCellValueToString(Cell cell) {
		if (cell == null) {
			return null;
		}
		String returnValue = null;
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC: // 数字
			Double doubleValue = cell.getNumericCellValue();
			// 格式化科学计数法，取一位整数
			DecimalFormat df = new DecimalFormat("0.00");
			returnValue = df.format(doubleValue);
			break;
		case Cell.CELL_TYPE_STRING: // 字符串
			returnValue = cell.getStringCellValue();
			break;
		case Cell.CELL_TYPE_BOOLEAN: // 布尔
			Boolean booleanValue = cell.getBooleanCellValue();
			returnValue = booleanValue.toString();
			break;
		case Cell.CELL_TYPE_BLANK: // 空值
			break;
		case Cell.CELL_TYPE_FORMULA: // 公式
			returnValue = cell.getCellFormula();
			break;
		case Cell.CELL_TYPE_ERROR: // 故障
			break;
		default:
			break;
		}
		return returnValue;
	}

	private static String convertCellValueToString(Cell cell, int type) {
		if (cell == null) {
			return null;
		}
		String returnValue = null;
		if(cell.getCellType()==Cell.CELL_TYPE_BLANK) {
			
		}else {
			switch (type) {
			case Cell.CELL_TYPE_NUMERIC: // 数字
				Double doubleValue = cell.getNumericCellValue();
				// 格式化科学计数法，取一位整数
				DecimalFormat df = new DecimalFormat("0.0");
				returnValue = df.format(doubleValue);
				break;
			case Cell.CELL_TYPE_STRING: // 字符串
				cell.setCellType(Cell.CELL_TYPE_STRING);
				returnValue = cell.getStringCellValue();
				break;
			case Cell.CELL_TYPE_BOOLEAN: // 布尔
				Boolean booleanValue = cell.getBooleanCellValue();
				returnValue = booleanValue.toString();
				break;
			case Cell.CELL_TYPE_BLANK: // 空值
				break;
			case Cell.CELL_TYPE_FORMULA: // 公式
				returnValue = cell.getCellFormula();
				break;
			case Cell.CELL_TYPE_ERROR: // 故障
				break;
			default:
				break;
			}
		}	
		return returnValue;
	}

	/**
	 * 提取每一行中需要的数据，构造成为一个结果数据对象
	 *
	 * 当该行中有单元格的数据为空或不合法时，忽略该行的数据 焊点信息
	 * 
	 * @param row       行数据
	 * @param boradlist
	 * @return 解析后的行数据对象，行数据错误时返回null
	 */
	private static WeldPointBoardInformation convertRowToData(Row row, List<BoardInformation> boradlist) {
		WeldPointBoardInformation resultData = new WeldPointBoardInformation();

		Cell cell;
		// 焊点编号
		cell = row.getCell(0);
		String weldno = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
		resultData.setWeldno(weldno);

		if (weldno == null || weldno.isEmpty()) {
			return null;
		}
		// 重要度
		cell = row.getCell(7);
		String importance = convertCellValueToString(cell);
		resultData.setImportance(importance);
		// 板组1
		cell = row.getCell(8);
		String partNo1 = convertCellValueToString(cell);
		// 板组2
		cell = row.getCell(11);
		String partNo2 = convertCellValueToString(cell);
		// 板组3
		cell = row.getCell(14);
		String partNo3 = convertCellValueToString(cell);
		
		resultData.setPartNo1(partNo1);
		resultData.setPartNo2(partNo2);
		resultData.setPartNo3(partNo3);

		// 根据板组信息通过焊点获取零件名称和板组编号、强度和GA/GI属性
		String boardnumber1 = ""; // 板材编号
		String boardname1 = ""; // 板材名称
		String partmaterial1 = ""; // 板材材质
		String partthickness1 = ""; // 板材板厚
		String sheetstrength1 = ""; // 材料强度(Mpa)
		String gagi1 = ""; // GA /GI

		String boardnumber2 = ""; // 板材编号
		String boardname2 = ""; // 板材名称
		String partmaterial2 = ""; // 板材材质
		String partthickness2 = ""; // 板材板厚
		String sheetstrength2 = ""; // 材料强度(Mpa)
		String gagi2 = ""; // GA /GI

		String boardnumber3 = ""; // 板材编号
		String boardname3 = ""; // 板材名称
		String partmaterial3 = ""; // 板材材质
		String partthickness3 = ""; // 板材板厚
		String sheetstrength3 = ""; // 材料强度(Mpa)
		String gagi3 = ""; // GA /GI

		if (boradlist.size() > 0) {
			for (int i = 0; i < boradlist.size(); i++) {
				BoardInformation bdf = boradlist.get(i);
				if (bdf != null && bdf.getPartn() != null) {
					if (partNo1 != null && !partNo1.isEmpty()) {
						if (bdf.getPartn().equals(partNo1)) {
							boardnumber1 = bdf.getBoardnumber();
							boardname1 = bdf.getBoardname();
							partmaterial1 = bdf.getPartmaterial();
							partthickness1 = bdf.getPartthickness();
							sheetstrength1 = bdf.getSheetstrength();
							gagi1 = bdf.getGagi();
						}
					}
					if (partNo2 != null && !partNo2.isEmpty()) {
						if (bdf.getPartn().equals(partNo2)) {
							boardnumber2 = bdf.getBoardnumber();
							boardname2 = bdf.getBoardname();
							partmaterial2 = bdf.getPartmaterial();
							partthickness2 = bdf.getPartthickness();
							sheetstrength2 = bdf.getSheetstrength();
							gagi2 = bdf.getGagi();
						}
					}
					if (partNo3 != null && !partNo3.isEmpty()) {
						if (bdf.getPartn().equals(partNo3)) {
							boardnumber3 = bdf.getBoardnumber();
							boardname3 = bdf.getBoardname();
							partmaterial3 = bdf.getPartmaterial();
							partthickness3 = bdf.getPartthickness();
							sheetstrength3 = bdf.getSheetstrength();
							gagi3 = bdf.getGagi();
						}
					}
				}

			}
		}
		// 如果板厚不是数值，则默认给空，避免报错
		if (!Util.isNumber(partthickness1)) {
			partthickness1 = "";
		}
		if (!Util.isNumber(partthickness2)) {
			partthickness2 = "";
		}
		if (!Util.isNumber(partthickness3)) {
			partthickness3 = "";
		}
		//获取厚薄板
		ArrayList hbboradlist = new ArrayList();
		if (partthickness1.isEmpty()&&!Util.isNumber(partthickness1)) {
			if(partNo1!=null && !partNo1.isEmpty()) {
				hbboradlist.add(partNo1);
			}
		}
		if (partthickness2.isEmpty()&&!Util.isNumber(partthickness2)) {
			if(partNo2!=null && !partNo2.isEmpty()) {
				hbboradlist.add(partNo2);
			}
		}
		if (partthickness2.isEmpty()&&!Util.isNumber(partthickness3)) {
			if(partNo3!=null && !partNo3.isEmpty()) {
				hbboradlist.add(partNo3);
			}
		}
		resultData.setHbborad(hbboradlist);
		// 板组1
		resultData.setBoardnumber1(boardnumber1);
		resultData.setBoardname1(boardname1);
		resultData.setPartmaterial1(partmaterial1);
		resultData.setPartthickness1(partthickness1);
		resultData.setStrength1(sheetstrength1);
		resultData.setGagi1(gagi1);

		// 板组2
		resultData.setBoardnumber2(boardnumber2);
		resultData.setBoardname2(boardname2);
		resultData.setPartmaterial2(partmaterial2);
		resultData.setPartthickness2(partthickness2);
		resultData.setStrength2(sheetstrength2);
		resultData.setGagi2(gagi2);

		// 板组3
		resultData.setBoardnumber3(boardnumber3);
		resultData.setBoardname3(boardname3);
		resultData.setPartmaterial3(partmaterial3);
		resultData.setPartthickness3(partthickness3);
		resultData.setStrength3(sheetstrength3);
		resultData.setGagi3(gagi3);

		int boradnum = 0;// 板层数
		if (!boardnumber1.isEmpty()) {
			boradnum++;
		}
		if (!boardnumber2.isEmpty()) {
			boradnum++;
		}
		if (!boardnumber3.isEmpty()) {
			boradnum++;
		}
		resultData.setLayersnum(Integer.toString(boradnum));// 板层数
		int ganum = 0;
		int ginum = 0;
		if (gagi1.equals("GA")) {
			ganum++;
		} else if (gagi1.equals("GI")) {
			ginum++;
		} else {

		}
		if (gagi2.equals("GA")) {
			ganum++;
		} else if (gagi2.equals("GI")) {
			ginum++;
		} else {

		}
		if (gagi3.equals("GA")) {
			ganum++;
		} else if (gagi3.equals("GI")) {
			ginum++;
		} else {

		}
		if (ganum == 0 && ginum == 0) {
			resultData.setGagi("");
		} else if (ganum != 0 && ginum == 0) {
			resultData.setGagi("GA");
		} else if (ganum == 0 && ginum != 0) {
			resultData.setGagi("GI");
		} else {
			resultData.setGagi("A/I"); // GA/GI
		}

		int num1 = 0;// 440
		int num2 = 0;// 590
		int num3 = 0;// >590
		if (sheetstrength1.isEmpty() && sheetstrength2.isEmpty() && sheetstrength3.isEmpty()) {
			num1 = 0;
			num2 = 0;
			num3 = 0;
		} else {
			if (sheetstrength1 != null && !sheetstrength1.isEmpty()) {
				int strength1 = (int) Double.parseDouble(sheetstrength1);
				if (strength1 == 440) {
					num1++;
				}
				if (strength1 == 590) {
					num2++;
				}
				if (strength1 > 590) {
					num3++;
				}
			}
			if (sheetstrength2 != null && !sheetstrength2.isEmpty()) {
				int strength2 = (int) Double.parseDouble(sheetstrength2);
				if (strength2 == 440) {
					num1++;
				}
				if (strength2 == 590) {
					num2++;
				}
				if (strength2 > 590) {
					num3++;
				}
			}
			if (sheetstrength3 != null && !sheetstrength3.isEmpty()) {
				int strength3 = (int) Double.parseDouble(sheetstrength3);
				if (strength3 == 440) {
					num1++;
				}
				if (strength3 == 590) {
					num2++;
				}
				if (strength3 > 590) {
					num3++;
				}
			}

		}
		// 板材强度
		resultData.setSheetstrength440(Integer.toString(num1));
		resultData.setSheetstrength590(Integer.toString(num2));
		resultData.setSheetstrength(Integer.toString(num3));

		// 基准板厚
		String basethickness = "";
		// 3层板取平均值四舍五入
		if (boradnum == 3) {
			if ((partthickness1 != null && !partthickness1.isEmpty())
					&& (partthickness2 != null && !partthickness2.isEmpty())
					&& (partthickness3 != null && !partthickness3.isEmpty())) {
				BigDecimal path1 = new BigDecimal(partthickness1);
				BigDecimal path2 = new BigDecimal(partthickness2);
				BigDecimal path3 = new BigDecimal(partthickness3);
				BigDecimal totalsum = path1.add(path2).add(path3);	
				String basenum = totalsum.divide(new BigDecimal("3"),6).toString();
//				double totalsum = Double.parseDouble(partthickness1) + Double.parseDouble(partthickness2)
//						+ Double.parseDouble(partthickness3);
//				double basenum = totalsum / 3;
				BigDecimal bd = new BigDecimal(basenum);
				BigDecimal bdvalue = bd.setScale(1, BigDecimal.ROUND_HALF_UP);
				basethickness = bdvalue.toString();
			}
		} else if (boradnum == 2) { // 2层板取薄板
			if (partthickness1 == null || partthickness1.isEmpty()) {
				if ((partthickness2 != null && !partthickness2.isEmpty())
						&& (partthickness3 != null && !partthickness3.isEmpty())) {

					if (Double.parseDouble(partthickness2) > Double.parseDouble(partthickness3)) {
						BigDecimal bd = new BigDecimal(Double.parseDouble(partthickness3));
						BigDecimal bdvalue = bd.setScale(1, BigDecimal.ROUND_HALF_UP);
						basethickness = bdvalue.toString();
					} else {
						BigDecimal bd = new BigDecimal(Double.parseDouble(partthickness2));
						BigDecimal bdvalue = bd.setScale(1, BigDecimal.ROUND_HALF_UP);
						basethickness = bdvalue.toString();
					}
				}
			} else if (partthickness2 == null || partthickness2.isEmpty()) {
				if ((partthickness1 != null && !partthickness1.isEmpty())
						&& (partthickness3 != null && !partthickness3.isEmpty())) {
					if (Double.parseDouble(partthickness1) > Double.parseDouble(partthickness3)) {
						BigDecimal bd = new BigDecimal(Double.parseDouble(partthickness3));
						BigDecimal bdvalue = bd.setScale(1, BigDecimal.ROUND_HALF_UP);
						basethickness = bdvalue.toString();
					} else {
						BigDecimal bd = new BigDecimal(Double.parseDouble(partthickness1));
						BigDecimal bdvalue = bd.setScale(1, BigDecimal.ROUND_HALF_UP);
						basethickness = bdvalue.toString();
					}
				}
			} else {
				if ((partthickness1 != null && !partthickness1.isEmpty())
						&& (partthickness2 != null && !partthickness2.isEmpty())) {
					if (Double.parseDouble(partthickness1) > Double.parseDouble(partthickness2)) {
						BigDecimal bd = new BigDecimal(Double.parseDouble(partthickness2));
						BigDecimal bdvalue = bd.setScale(1, BigDecimal.ROUND_HALF_UP);
						basethickness = bdvalue.toString();
					} else {
						BigDecimal bd = new BigDecimal(Double.parseDouble(partthickness1));
						BigDecimal bdvalue = bd.setScale(1, BigDecimal.ROUND_HALF_UP);
						basethickness = bdvalue.toString();
					}
				}
			}
		} else if (boradnum == 1) {
			if (partthickness1 != null && !partthickness1.isEmpty()) {
				BigDecimal bd = new BigDecimal(Double.parseDouble(partthickness1));
				BigDecimal bdvalue = bd.setScale(1, BigDecimal.ROUND_HALF_UP);
				basethickness = bdvalue.toString();
			} else if (partthickness2 != null && !partthickness2.isEmpty()) {
				BigDecimal bd = new BigDecimal(Double.parseDouble(partthickness2));
				BigDecimal bdvalue = bd.setScale(1, BigDecimal.ROUND_HALF_UP);
				basethickness = bdvalue.toString();
			} else if (partthickness3 != null && !partthickness3.isEmpty()){
				BigDecimal bd = new BigDecimal(Double.parseDouble(partthickness3));
				BigDecimal bdvalue = bd.setScale(1, BigDecimal.ROUND_HALF_UP);
				basethickness = bdvalue.toString();
			}else {
				basethickness = "";
			}
		} else {
		}		
		// 1.2g:只要焊点关联的板材中"-"后面部分含有1180数字的就是1.2g高强才
		boolean flag = false;

		if (partmaterial1 != null && !partmaterial1.isEmpty()) {
			String Sheetstrength = "";
			String[] str = partmaterial1.split("-");
			if (str.length > 1) {
				String tempstr = str[1].trim();
				if (tempstr != null && !"".equals(tempstr)) {
					for (int K = 0; K < tempstr.length(); K++) {
						if (tempstr.charAt(K) >= 48 && tempstr.charAt(K) <= 57) {
							Sheetstrength += tempstr.charAt(K);
						}
					}
				}
			}
			if (Sheetstrength.equals("1180")) {
				flag = true;
			}
		}
		if (partmaterial2 != null && !partmaterial2.isEmpty()) {
			String Sheetstrength = "";
			String[] str = partmaterial2.split("-");
			if (str.length > 1) {
				String tempstr = str[1].trim();
				if (tempstr != null && !"".equals(tempstr)) {
					for (int K = 0; K < tempstr.length(); K++) {
						if (tempstr.charAt(K) >= 48 && tempstr.charAt(K) <= 57) {
							Sheetstrength += tempstr.charAt(K);
						}
					}
				}
			}
			if (Sheetstrength.equals("1180")) {
				flag = true;
			}
		}
		if (partmaterial3 != null && !partmaterial3.isEmpty()) {
			String Sheetstrength = "";
			String[] str = partmaterial3.split("-");
			if (str.length > 1) {
				String tempstr = str[1].trim();
				if (tempstr != null && !"".equals(tempstr)) {
					for (int K = 0; K < tempstr.length(); K++) {
						if (tempstr.charAt(K) >= 48 && tempstr.charAt(K) <= 57) {
							Sheetstrength += tempstr.charAt(K);
						}
					}
				}
			}
			if (Sheetstrength.equals("1180")) {
				flag = true;
			}
		}
		if (flag) {
			resultData.setSheetstrength12("1.2g");
			basethickness = getMinnum(partthickness1, partthickness2, partthickness3);
		} else {
			resultData.setSheetstrength12("g");
		}
		resultData.setBasethickness(basethickness);
		
		return resultData;
	}

	/**
	 * 提取每一行中需要的数据，构造成为一个结果数据对象
	 *
	 * 当该行中有单元格的数据为空或不合法时，忽略该行的数据 板组信息
	 * 
	 * @param row 行数据
	 * @return 解析后的行数据对象，行数据错误时返回null
	 */
	private static BoardInformation convertRowToBoradData(Row row) {
		BoardInformation resultData = new BoardInformation();

		Cell cell;

		// 序号
		cell = row.getCell(0);
		String rowNum = convertCellValueToString(cell);
		if (rowNum == null) {
			rowNum = "";
		}
		resultData.setRowNum(rowNum);

		// 板组编号
		cell = row.getCell(1);
		String bzcode = convertCellValueToString(cell);
		if (bzcode == null) {
			bzcode = "";
		}
		resultData.setBoardnumber(bzcode);

		// 零 件 号
		cell = row.getCell(2);
		String partn = convertCellValueToString(cell);
		if (partn == null || partn.isEmpty()) {
			return null;
		}
		resultData.setPartn(partn);
		// 零 件 名 称
		cell = row.getCell(4);
		String boardname = convertCellValueToString(cell);
		if (boardname == null) {
			boardname = "";
		}
		resultData.setBoardname(boardname);

		// 获取材 质
		cell = row.getCell(5);
		String partmaterial = convertCellValueToString(cell);
		if (partmaterial == null) {
			partmaterial = "";
		}
		resultData.setPartmaterial(partmaterial);

		// 获取板 厚
		cell = row.getCell(6);
		String partthickness = convertCellValueToString(cell,Cell.CELL_TYPE_STRING);
		if (partthickness == null) {
			partthickness = "";
		}
		resultData.setPartthickness(partthickness);

		// 获取板 厚单位
		cell = row.getCell(7);
		String maunit = convertCellValueToString(cell);
		if (maunit == null) {
			maunit = "";
		}
		resultData.setMaunit(maunit);

		// 获取强 度
		cell = row.getCell(8);
		String sheetstrength = convertCellValueToString(cell,Cell.CELL_TYPE_STRING);
		if (sheetstrength == null) {
			sheetstrength = "";
		}
		resultData.setSheetstrength(sheetstrength);

		// 获取强 度单位
		cell = row.getCell(9);
		String thunit = convertCellValueToString(cell);
		if (thunit == null) {
			thunit = "";
		}
		resultData.setThunit(thunit);

		// 获取GA/GI
		cell = row.getCell(10);
		String gagi = convertCellValueToString(cell);
		if (gagi == null) {
			gagi = "";
		}
		resultData.setGagi(gagi);

		return resultData;
	}

	private static List<WeldPointInfo> convertRowToWeldData(Row row) {
		List<WeldPointInfo> resultData = new ArrayList<WeldPointInfo>();
		Cell cell;
		// 板件1
		cell = row.getCell(8);
		String partno1 = convertCellValueToString(cell);
		if (partno1 != null && !partno1.isEmpty()) {
			WeldPointInfo weld = new WeldPointInfo();
			weld.setPartno(partno1);
			cell = row.getCell(10);
			String partmaterial = convertCellValueToString(cell);
			weld.setPartmaterial(partmaterial);
			cell = row.getCell(9);
			String partthickness = convertCellValueToString(cell);
			weld.setPartthickness(partthickness);
			resultData.add(weld);
		}
		// 板件2
		cell = row.getCell(11);
		String partno2 = convertCellValueToString(cell);
		if (partno2 != null && !partno2.isEmpty()) {
			WeldPointInfo weld = new WeldPointInfo();
			weld.setPartno(partno2);
			cell = row.getCell(13);
			String partmaterial = convertCellValueToString(cell);
			weld.setPartmaterial(partmaterial);
			cell = row.getCell(12);
			String partthickness = convertCellValueToString(cell);
			weld.setPartthickness(partthickness);
			resultData.add(weld);
		}
		// 板件3
		cell = row.getCell(14);
		String partno3 = convertCellValueToString(cell);
		if (partno3 != null && !partno3.isEmpty()) {
			WeldPointInfo weld = new WeldPointInfo();
			weld.setPartno(partno3);
			cell = row.getCell(16);
			String partmaterial = convertCellValueToString(cell);
			weld.setPartmaterial(partmaterial);
			cell = row.getCell(15);
			String partthickness = convertCellValueToString(cell);
			weld.setPartthickness(partthickness);
			resultData.add(weld);
		}

		return resultData;
	}

	/*
	 * 取最小值
	 */
	private static String getMinnum(String str1, String str2, String str3) {
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
		if (minstr.equals("9999")) {
			minstr = "";
		}
		if(!minstr.isEmpty()) {
			BigDecimal bd = new BigDecimal(Double.parseDouble(minstr));
			BigDecimal bdvalue = bd.setScale(1, BigDecimal.ROUND_HALF_UP);
			minstr = bdvalue.toString();
		}	
		return minstr;
	}

	/**
	 * 获取材料对照表
	 * @param app
	 * @param prefrencename
	 * @return
	 */
	public static Map<String,List<String>> getMaterialComparisonTable(AbstractAIFUIApplication app,String prefrencename) {
		Map<String,List<String>> listmap = new HashMap<>();
		try {

			File file = null;
			Workbook workbook = null;
			TCSession session = (TCSession) app.getSession();
			TCPreferenceService preferenceService = session.getPreferenceService();

//			String str = preferenceService.getPreferenceDescription(prefrencename);
//			if (str != null) {
//				String value = preferenceService.getStringValue(prefrencename);
				if (prefrencename != null) {
					TCComponentDatasetType datatype = (TCComponentDatasetType) session.getTypeComponent("Dataset");
					TCComponentDataset dataset = datatype.find(prefrencename);
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
								listmap = parseMaterialExcelForpara(workbook);
							}
							if (type.equals("MSExcelX")) {
								workbook = new XSSFWorkbook(inputStream);
								listmap = parseMaterialExcelForpara(workbook);
							}
						}
					}
				}
			
			return listmap;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return listmap;
	}

	private static Map<String, List<String>> parseMaterialExcelForpara(Workbook workbook) {
		// TODO Auto-generated method stub
		try {
			Map<String, List<String>> map = new HashMap<>();
			Sheet sheet = workbook.getSheetAt(0);
			// 校验sheet是否合法
			if (sheet == null) {
				return null;
			}
			for(int i=1;i<sheet.getPhysicalNumberOfRows();i++)
			{
				Row row = sheet.getRow(i);
				String MaterialNo = "";
				String gagi = "";
				String isCule = "";
				if(row!=null)
				{
					Cell cell = row.getCell(1);
					if(cell!=null)
					{
						MaterialNo = cell.getStringCellValue().trim().replace(" ", "");
					}
					Cell cell1 = row.getCell(2);
					if(cell1!=null)
					{
						gagi = cell1.getStringCellValue().trim().replace(" ", "");
					}
					Cell cell2 = row.getCell(3);
					if(cell2!=null)
					{
						isCule = cell2.getStringCellValue().trim().replace(" ", "");
					}							
				}
				if(!MaterialNo.trim().isEmpty())
				{
					List<String> list = new ArrayList<String>();
					list.add(gagi.trim());
					list.add(isCule.trim());
					map.put(MaterialNo.trim(), list);
				}
				
			}
			return map;
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null;
	}

}
