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

		//��ֺ������������������ϵ
		//����ֺ���ĵ�����汾��DFL9MEDocumentRevision��ͨ����ϵWH7_WorkSheet_Rel��ѡ���ԭʼ�����ĵ��汾��DFL9MEDocumentRevision������
		addRel(docRev, spliteDocRev, "WH7_WorkSheet_Rel");
		//��ֺ���빤����������ϵ
		//1��ͨ������Doc��ȡ��λ����B8_BIWMEProcStatRevision

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
			System.out.println(">>>WARN:δ�ҵ��������չ����Ĺ�λ���գ�������");
			return;
		}


		TCComponent[] childrens = processStation.getReferenceListProperty("ps_children");
		if(childrens==null	)
		{
			System.out.println(">>>WARN:��λ�����¼�����Ϊ�գ�������");
			return;
		}

		//2����ȡ��λ���յ��¼��������汾B8_BIWDiscreteOPRevision��B8_BIWOperationRevision��B8_BIWArcWeldOPRevision
		String docName = Util.getProperty(spliteDocRev, "object_name");
		System.out.println("spliteDocRev name:"+docName);

		//"�㺸-RSW"
		docName = docName.replace("�㺸-RSW", "�㺸RSW");
		docName = docName.replace("�㺸-PSW", "�㺸PSW");

		docName = docName.substring(docName.lastIndexOf("-")+1, docName.length());//ȡ���һ��"-"��Ĳ���
		docName = docName.replaceFirst("\\d+","");//ȥ����ͷ������

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

				//				object_type	object_name	object_name��*��ʾͨ�����
				//				B8_BIWDiscreteOPRevision  	= R*��
				//				�磺R1001\IRCN-4040385	*�㺸-RSW��
				//				�����ĵ��汾������ͨ��IMAN_specification���ص�MSExcelX�ļ�,��ȡ��Ԫ��CM10��ֵȥƥ�乤�����Ƶ�һ����\�����ֵ����IRCN-4040385��,ƥ�����򽫶�Ӧ���ĵ�����汾���ص���ǰ����汾��
				//					!= R*��
				//				�磺W1-H1-GW006\UXH-C1395-ZC	*�㺸-PSW ��
				//				�����ĵ��汾������ͨ��IMAN_specification���ص�MSExcelX�ļ�,��ȡ��Ԫ��CM10��ֵȥƥ�乤�����Ƶ�һ����\�����ֵ����UXH-C1395-ZC��,ƥ�����򽫶�Ӧ���ĵ�����汾���ص���ǰ����汾��
				//String opName2 = opName.substring(opName.indexOf("\\")+1,opName.length());




				System.out.println(">>>opName2:"+opName);
				if(opName.startsWith("R")&&docName.startsWith("�㺸RSW"))
				{
					if(!opName.contains("\\")||opName.endsWith("\\"))
					{
						//ֻ��ȡCG10() CG12()��ֵ��ȥ��\�������Ƶ�ֵ�Ƚ�
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
				}else if(!opName.startsWith("R")&&docName.startsWith("�㺸PSW"))
				{
					if(!opName.contains("\\")||opName.endsWith("\\"))
					{
						//ֻ��ȡCG10() CG12()��ֵ��ȥ��\�������Ƶ�ֵ�Ƚ�
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
				if("�ϼ�".equals(opName)&&docName.equals("���ɱ�"))
				{
					addRel = true;
				}else if("�ϼ�".equals(opName)&&docName.equals("����ͼ"))
				{
					addRel = true;
				} 
				else if(opName.equals(docName))
				{
					addRel = true;
				}
				/*
				 * else if("��í".equals(opName)&&docName.contains("��í")) { addRel = true; }else
				 * if("í��".equals(opName)&&docName.contains("í��")) { addRel = true; }else
				 * if("��˨��".equals(opName)&&docName.contains("��˨��")) { addRel = true; }else
				 * if("������".equals(opName)&&docName.contains("������")) { addRel = true; }else
				 * if("��ĸ��".equals(opName)&&docName.contains("��ĸ��")) { addRel = true; }else
				 * if("�ɽ���ĥ".equals(opName)&&docName.contains("�ɽ���ĥ")) { addRel = true; }else
				 * if("װ��".equals(opName)&&docName.contains("װ��")) { addRel = true; }else
				 * if("�������".equals(opName)&&docName.contains("�������")) { addRel = true; }else
				 * if("���".equals(opName)&&docName.contains("���")) { addRel = true; }
				 */
			}else if("B8_BIWArcWeldOPRevision".equals(objectType))
			{
				if("��������".equals(opName)&&docName.equals("������ҵ"))
				{
					addRel = true;
				}else if("���⺸����".equals(opName)&&docName.equals("���⺸"))
				{
					addRel = true;
				}else if("ǥ������".equals(opName)&&docName.equals("ǥ��"))
				{
					addRel = true;
				}else if("Ϳ������".equals(opName)&&docName.equals("Ϳ��"))
				{
					addRel = true;
				}else if("ARPLAS����".equals(opName)&&docName.equals("ARPLAS��"))
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
				System.out.println(">>>>>>ƥ��Ĺ�������ҵ��"+opName+"--"+docName);

			}else
			{
				//				System.out.println("WARN:δƥ�䣺"+opName+"--"+docName);
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
			System.out.println(">>>WARN:δ�ҵ��������չ����Ĺ�λ���գ�������");
			return null;
		}

		TCComponentItemRevision rev = (TCComponentItemRevision) refComps[0].getComponent();
		TCComponentItemRevision processStat = rev.getItem().getLatestItemRevision();//ȡ���°汾
		System.out.println(">>>��װ��λ���գ�"+processStat);

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
					System.out.println( "ERROR:Sheet������"); 
					return result;
				}

				for (int[] cellNums :cellsList) {
					Row row = sheet.getRow(cellNums[0]);
					Row row2 = sheet.getRow(cellNums[2]);

					if(row==null)
					{
						System.out.println( "ERROR:Sheet��"+cellNums[0]+"��Ϊ��"); 
						return result;
					}

					if(row2==null)
					{
						System.out.println( "ERROR:Sheet��"+cellNums[2]+"��Ϊ��"); 
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
			System.out.println( "ERROR:Excel�ļ�������"); 
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
					System.out.println( "ERROR:Sheet������"); 
					return "";
				}

				Row row = sheet.getRow(rowNum);
				if(row==null)
				{
					System.out.println( "ERROR:Sheet��"+rowNum+"��Ϊ��"); 
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
			System.out.println( "ERROR:Excel�ļ�������"); 
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
