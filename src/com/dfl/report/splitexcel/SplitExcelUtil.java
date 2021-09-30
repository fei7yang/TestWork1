package com.dfl.report.splitexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.dfl.report.util.ReportUtils;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCException;

/**
 * 
 * @author huangbin
 * @since 2021/8/10
 *
 */
public class SplitExcelUtil {

	private static FileInputStream inputStream = null;

	public static void setDocRelation(TCComponentItemRevision processStation, TCComponentItemRevision docRev,TCComponentItemRevision spliteDocRev, File newFile) throws TCException
	{
		inputStream = null;
		//CG10\CM10,CG10\CM11, CG12\CM12, CG12\CM13
		List<int[]> cellNumList = new ArrayList<int[]>();
		cellNumList.add(new int[] {9,84,9,90});
		cellNumList.add(new int[] {9,84,10,90});
		cellNumList.add(new int[] {11,84,11,90});
		cellNumList.add(new int[] {11,84,12,90});

		//拆分后的与整本建立关联关系
		//将拆分后的文档对象版本（DFL9MEDocumentRevision）通过关系WH7_WorkSheet_Rel与选择的原始工艺文档版本（DFL9MEDocumentRevision）关联
		addRel(docRev, spliteDocRev, "WH7_WorkSheet_Rel");
		//拆分后的与工序建立关联关系
		//1、通过整本Doc获取工位工艺B8_BIWMEProcStatRevision

		//2021/08/25
//		BU51=(50,72) 						dfl9_wh7WorkName
//		BU49=(48,72) 						dfl9_wh7WorkSheetName
//		CQ52 DD51/DI51 =(51,94) (50,107)/(50,112)		dfl9_wh7OperationNo
//		DE49(48,108)							dfl9_wh7VersionNo
		String BU51 = getExcelValue(newFile,50,72,"BU51");

		String BU49 = getExcelValue(newFile,48,72,"BU51");

		String CQ52 = getExcelValue(newFile,51,94,"CQ52");
		String DD51 = getExcelValue(newFile,50,107,"DD51");
		String DI51 = getExcelValue(newFile,50,112,"DI51");

		String DE49 = getExcelValue(newFile,48,108,"DE49");

		spliteDocRev.setProperty("dfl9_wh7WorkName", BU51);
		spliteDocRev.setProperty("dfl9_wh7WorkSheetName", BU49);
		spliteDocRev.setProperty("dfl9_wh7OperationNo", CQ52+" "+DD51+"/"+DI51);
		spliteDocRev.setProperty("dfl9_wh7Phase", DE49);
		
		if(processStation==null)
		{
			System.out.println(">>>WARN:未找到整本工艺关联的工位工艺，跳过！");
			return;
		}


		TCComponent[] childrens = processStation.getReferenceListProperty("ps_children");
		if(childrens==null	)
		{
			System.out.println(">>>WARN:工位工艺下级工序为空，跳过！");
			return;
		}

		//2、获取工位工艺的下级工序对象版本B8_BIWDiscreteOPRevision、B8_BIWOperationRevision、B8_BIWArcWeldOPRevision
		String docName = Util.getProperty(spliteDocRev, "object_name");
		System.out.println("spliteDocRev name:"+docName);

		//"点焊-RSW"
		docName = docName.replace("点焊-RSW", "点焊RSW");
		docName = docName.replace("点焊-PSW", "点焊PSW");

		docName = docName.substring(docName.lastIndexOf("-")+1, docName.length());//取最后一个"-"后的部分
		docName = docName.replaceFirst("\\d+","");//去掉开头的数字

		System.out.println("sheet name:"+docName);

		int opCount = childrens.length ;
		for (int i = 0; i < opCount; i++) 
		{
			TCComponent opRev = childrens[i];
			String objectType = opRev.getType();
			String opName = Util.getProperty(opRev, "object_name");

			boolean addRel = false;
			if("B8_BIWDiscreteOPRevision".equals(objectType))
			{

				//				object_type	object_name	object_name（*表示通配符）
				//				B8_BIWDiscreteOPRevision  	= R*，
				//				如：R1001\IRCN-4040385	*点焊-RSW：
				//				解析文档版本对象下通过IMAN_specification挂载的MSExcelX文件,获取单元格CM10的值去匹配工序名称第一个“\”后的值（如IRCN-4040385）,匹配上则将对应的文档对象版本挂载到当前工序版本下
				//					!= R*，
				//				如：W1-H1-GW006\UXH-C1395-ZC	*点焊-PSW ：
				//				解析文档版本对象下通过IMAN_specification挂载的MSExcelX文件,获取单元格CM10的值去匹配工序名称第一个“\”后的值（如UXH-C1395-ZC）,匹配上则将对应的文档对象版本挂载到当前工序版本下
				//String opName2 = opName.substring(opName.indexOf("\\")+1,opName.length());




				System.out.println(">>>opName2:"+opName);
				if(opName.startsWith("R")&&docName.startsWith("点焊RSW"))
				{
					if(!opName.contains("\\")||opName.endsWith("\\"))
					{
						//只需取CG10() CG12()的值与去掉\后工序名称的值比较
						String name = opName.replace("\\", "");
						String cg10 = getExcelValue(newFile,9,84,"CG10");
						String cg12 = getExcelValue(newFile,11,84,"CG12");
						if(name.equals(cg10)||name.equals(cg12))
						{
							addRel = true;
						}
					}else
					{
						List<String> resultList =  SplitExcelUtil.getExcelValues(newFile,cellNumList);//CM10 9,90
						System.out.println("resultList:"+resultList.toString());
						if(resultList.contains(opName))
						{
							addRel = true;
						}
					}
					//					String cm10Value = getExcelValue(newFile);
					//					if(cm10Value.equals(opName2))
					//					{
					//						addRel = true;
					//					}
				}else if(!opName.startsWith("R")&&docName.startsWith("点焊PSW"))
				{
					if(!opName.contains("\\")||opName.endsWith("\\"))
					{
						//只需取CG10() CG12()的值与去掉\后工序名称的值比较
						String name = opName.replace("\\", "");
						String cg10 = getExcelValue(newFile,9,84,"CG10");
						String cg12 = getExcelValue(newFile,11,84,"CG12");
						if(name.equals(cg10)||name.equals(cg12))
						{
							addRel = true;
						}
					}else
					{
						List<String> resultList =  SplitExcelUtil.getExcelValues(newFile,cellNumList);//CM10 9,90
						System.out.println("resultList:"+resultList.toString());
						if(resultList.contains(opName))
						{
							addRel = true;
						}
					}
				}

			}else if("B8_BIWOperationRevision".equals(objectType))
			{
				if("上件".equals(opName)&&docName.equals("构成表"))
				{
					addRel = true;
				}else if("上件".equals(opName)&&docName.equals("构成图"))
				{
					addRel = true;
				} 
				else if(opName.equals(docName))
				{
					addRel = true;
				}
				/*
				 * else if("拉铆".equals(opName)&&docName.contains("拉铆")) { addRel = true; }else
				 * if("铆接".equals(opName)&&docName.contains("铆接")) { addRel = true; }else
				 * if("螺栓焊".equals(opName)&&docName.contains("螺栓焊")) { addRel = true; }else
				 * if("螺柱焊".equals(opName)&&docName.contains("螺柱焊")) { addRel = true; }else
				 * if("螺母焊".equals(opName)&&docName.contains("螺母焊")) { addRel = true; }else
				 * if("飞溅打磨".equals(opName)&&docName.contains("飞溅打磨")) { addRel = true; }else
				 * if("装配".equals(opName)&&docName.contains("装配")) { addRel = true; }else
				 * if("建付规格".equals(opName)&&docName.contains("建付规格")) { addRel = true; }else
				 * if("检查".equals(opName)&&docName.contains("检查")) { addRel = true; }
				 */
			}else if("B8_BIWArcWeldOPRevision".equals(objectType))
			{
				if("弧焊工序".equals(opName)&&docName.equals("弧焊作业"))
				{
					addRel = true;
				}else if("激光焊工序".equals(opName)&&docName.equals("激光焊"))
				{
					addRel = true;
				}else if("钎焊工序".equals(opName)&&docName.equals("钎焊"))
				{
					addRel = true;
				}else if("涂胶工序".equals(opName)&&docName.equals("涂胶"))
				{
					addRel = true;
				}else if("ARPLAS工序".equals(opName)&&docName.equals("ARPLAS焊"))
				{
					addRel = true;
				}else if(opName.equals(docName))
				{
					addRel = true;
				}
			}	

			if(addRel)
			{
				addRel(opRev, spliteDocRev, "WH7_WorkSheet_Rel");
				System.out.println(">>>>>>匹配的工序与作业表："+opName+"--"+docName);

			}else
			{
				//				System.out.println("WARN:未匹配："+opName+"--"+docName);
			}
		}
			try {
				if(inputStream!=null)
				{
					inputStream.close();
				}
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
	}



	public static  TCComponentItemRevision getProcessStation(TCComponentItemRevision docRev) throws TCException {
		// TODO Auto-generated method stub
		AIFComponentContext[] refComps = docRev.getItem().whereReferencedByTypeRelation(new String[] {"B8_BIWMEProcStatRevision"}, new String[] {"IMAN_reference"});
		if(refComps==null||refComps.length<=0)
		{
			System.out.println(">>>WARN:未找到整本工艺关联的工位工艺，跳过！");
			return null;
		}

		TCComponentItemRevision rev = (TCComponentItemRevision) refComps[0].getComponent();
		TCComponentItemRevision processStat = rev.getItem().getLatestItemRevision();//取最新版本
		System.out.println(">>>焊装工位工艺："+processStat);

		return processStat;
	}

	public static Workbook workbook;
	public static List<String> getExcelValues(File xlsFile,List<int[]> cellsList) {
		// TODO Auto-generated method stub

		List<String> result = new ArrayList<String>();
		if(xlsFile!=null)
		{
			try {
				if(inputStream==null)
				{
					inputStream = new FileInputStream(xlsFile);
				    workbook = ReportUtils.getWorkbook(xlsFile.getAbsolutePath());
				}

				if(workbook==null)
				{
					System.out.println( "ERROR:workbook==null"); 
					return result;
				}
				Sheet sheet = workbook.getSheetAt(0);
				if (sheet == null)
				{
					System.out.println( "ERROR:Sheet不存在"); 
					return result;
				}

				for (int[] cellNums :cellsList) {
					Row row = sheet.getRow(cellNums[0]);
					Row row2 = sheet.getRow(cellNums[2]);

					if(row==null)
					{
						System.out.println( "ERROR:Sheet第"+cellNums[0]+"行为空"); 
						return result;
					}

					if(row2==null)
					{
						System.out.println( "ERROR:Sheet第"+cellNums[2]+"行为空"); 
						return result;
					}


					Cell cell11 = row.getCell(cellNums[1]);
					Cell cell12 = row2.getCell(cellNums[3]);

					//					Cell cell30 = row.getCell(30);//AE45O
					String value1 = ReportUtils.getCellValueString(cell11);
					String value2 = ReportUtils.getCellValueString(cell12);

					//					String value2 = ReportUtil.getCellValueString(cell30);
					if(value1!=null&&value2!=null)
					{
						System.out.println(cellNums[0]+","+cellNums[1]+ " value:"+value1); 
						result.add(value1+"\\"+value2);
					}
				}

			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
				System.out.println( e.getMessage());
				System.out.println( e);
			}
		}else
		{
			System.out.println( "ERROR:Excel文件不存在"); 
		}

		return result;
	}




	public static String getExcelValue(File xlsFile,int rowNum,int columnNum,String cellName) {
		// TODO Auto-generated method stub


		if(xlsFile!=null)
		{
			try {
				if(inputStream==null)
				{
					inputStream = new FileInputStream(xlsFile);
				    workbook = ReportUtils.getWorkbook(xlsFile.getAbsolutePath());
				}

				
				if(workbook==null)
				{
					System.out.println( "ERROR:workbook==null"); 
					return "";
				}
				Sheet sheet = workbook.getSheetAt(0);
				if (sheet == null)
				{
					System.out.println( "ERROR:Sheet不存在"); 
					return "";
				}

				Row row = sheet.getRow(rowNum);
				if(row==null)
				{
					System.out.println( "ERROR:Sheet第"+rowNum+"行为空"); 
					return "";
				}
				Cell cell12 = row.getCell(columnNum);//M47
				//				Cell cell30 = row.getCell(30);//AE45O
				String value1 = ReportUtils.getCellValueString(cell12);
				//				String value2 = ReportUtil.getCellValueString(cell30);
				if(value1!=null)
				{
					System.out.println(cellName+ " value:"+value1); 
					return value1;
				}


			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
				System.out.println( e.getMessage());
				System.out.println( e);
			}finally {
				
			}
		}else
		{
			System.out.println( "ERROR:Excel文件不存在"); 
		}
		return "";
	}




	public static void addRel(TCComponent pobj, TCComponent secObject, String relation) throws TCException {
		// TODO Auto-generated method stub
		TCComponent[] comps = pobj.getRelatedComponents(relation);
		for (TCComponent tcComponent : comps) {
			if(tcComponent.equals(secObject))
			{
				return;
			}
		}

		pobj.add(relation, secObject);


	}

	public static void deleteRel(TCComponent pobj, TCComponent secObject, String relation) throws TCException {
		// TODO Auto-generated method stub

		TCComponent[] comps = pobj.getRelatedComponents(relation);
		for (TCComponent tcComponent : comps) {
			if(tcComponent.equals(secObject))
			{
				pobj.remove(relation, secObject);
			}
		}

	}

}
