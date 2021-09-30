package com.dfl.report.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import javax.swing.JOptionPane;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dfl.report.common.TCComponentProjectUtils;
import com.dfl.report.common.TCComponentReleaseStatusUtils;
import com.dfl.report.common.TCComponentUtils;
import com.dfl.report.common.TCPreferenceServiceUtils;
import com.dfl.report.common.TCUtils;
import com.teamcenter.rac.kernel.DeepCopyInfo;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentItemType;
import com.teamcenter.rac.kernel.TCComponentProject;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.schemas.soa._2006_03.exceptions.ServiceException;
import com.teamcenter.services.rac.core._2008_06.DataManagement.CreateResponse;



public class ReportUtils {
	
	public static final String DFL_Project_VehicleNo = "DFL_Project_VehicleNo"; //$NON-NLS-1$
	public static final String IMAN_reference = "IMAN_reference"; //$NON-NLS-1$
	public static final String B8_BIWProcDocRevision = "B8_BIWProcDocRevision"; //$NON-NLS-1$
	public static final String B8_BIWProcDoc = "B8_BIWProcDoc"; //$NON-NLS-1$
	public static final String B8_BIWProcDocRevisionMaster = "B8_BIWProcDocRevisionMaster"; //$NON-NLS-1$
	public static final String DFL9MEDocument = "DFL9MEDocument"; //$NON-NLS-1$
	public static final String DFL9MEDocumentRevision = "DFL9MEDocumentRevision"; //$NON-NLS-1$
	public static final String DFL9MEDocumentRevisionMaster = "DFL9MEDocumentRevisionMaster"; //$NON-NLS-1$
	/**
	 * ��ȡ����������-��Ŀ��ѡ�� 
	 */
	public static Map<String,String> getDFL_Project_VehicleNo(){
		Map<String,String> projVehMap = new HashMap<String,String>();
		String[] proj_vehArr = null;
		try {
			proj_vehArr = TCPreferenceServiceUtils.getPrefernceValues(DFL_Project_VehicleNo, null);
		} catch (Exception e) {
			e.printStackTrace();
		}
		if(proj_vehArr != null){
			for(String proj_veh : proj_vehArr){
				if(proj_veh.contains(":")){ //$NON-NLS-1$
					String[] arr = proj_veh.split(":"); //$NON-NLS-1$
					if(arr.length == 2){
						projVehMap.put(arr[0], arr[1]);
					}
				}
			}
		}
		return projVehMap;
	}
	
	/**
	 * 
	 * @param productBOPList ѡ�е�productBOP����
	 * @param dfl9_process_type  
	 * @param dfl9_process_file_type
	 * @param isExist �Ƿ����MEDocument
	 * @param isGoon  �Ƿ����ִ��
	 * @return
	 * @throws TCException
	 * 1����ѡ���BOP�汾��ʹ�����ù�ϵ��������ֵ����MEDocument�汾����
	 * 2�����û���ҵ���������excel�ļ�����ΪMSExcelX���ݼ���ʹ�ù���ϵ���ص�һ���½�MEDocument�汾������Ӧ��MEDocument�汾����ʹ�����ù�ϵ���ص����е�ProductBOP�汾��
	 * 3������ҵ�����ProductBOP�汾���Ƴ�ԭ�ϰ汾MEDocument�汾����ʹ�����ù�ϵ������MEDocument�汾����  
	 * 		����ҵ�MEDocument�汾��������Ҫ�ж�MEDocument�汾�����Ƿ���working״̬�������С��ѷ��ţ�
	 * 		�������������Ҫ�������飬�ñ����Ѿ��������д�������
	 * 		������ѷ��ţ�����Ҫ��ʾ���ñ����ѷ��ţ��Ƿ����棿���û�ѡ���ǡ�����֮ǰ�������߼�����ѡ�񡰷��������
	 * 		�����working״̬������Ҫ��ʾ���ñ����Ѵ��ڣ��Ƿ񸲸ǲ����������û�ѡ���ǡ���ֱ�Ӹ���MEDocument�汾�µ����ݼ���ѡ�񡰷��������
	 */
	public static GenerateReportInfo  beforeGenerateReportAction(TCComponentItemRevision BOPItemRev,GenerateReportInfo info) throws TCException {
		
		//��ProductBOP�汾��ʹ�����ù�ϵ��������ֵ���ˣ�dfl9_process_type��Ӧֵ��Z����dfl9_process_file_type��Ӧֵ��FC����MEDocument�汾�������û���ҵ���ִ��2.4������ҵ���ִ��2.5
		TCComponentItemRevision b8MEDocumentRev = null;
			TCComponent[] relatedComps = TCComponentUtils.getCompsByRelation(BOPItemRev, ReportUtils.IMAN_reference);
			if(relatedComps != null && relatedComps.length > 0){
				for(TCComponent relatedComp : relatedComps){
					String type = relatedComp.getType();
					if(relatedComp instanceof TCComponentItem && ReportUtils.DFL9MEDocument.equals(type)){
						TCComponentItem itemdocument = (TCComponentItem)relatedComp;
						TCComponentItemRevision document = itemdocument.getLatestItemRevision();
						String process_type = document.getStringProperty("dfl9_process_type"); //$NON-NLS-1$
						String process_file_type = document.getStringProperty("dfl9_process_file_type"); //$NON-NLS-1$
						String documentname = document.getStringProperty("object_name");
						if(info.isFlag()) {
							if(info.getDFL9_process_type().equals(process_type) && info.getDFL9_process_file_type().equals(process_file_type)){
								//�ҵ�
								info.setExist(true);
								b8MEDocumentRev = document;
								break;
							}
						}else {
							if(info.getDFL9_process_type().equals(process_type) && info.getDFL9_process_file_type().equals(process_file_type)&&info.getmeDocumentName().equals(documentname)){
								//�ҵ�
								info.setExist(true);
								b8MEDocumentRev = document;
								break;
							}
						}
						
					}
				}
			}	
			//�ҵ�
			if(info.isExist()){
				//����ҵ�MEDocument�汾��������Ҫ�ж�MEDocument�汾�����Ƿ���working״̬�������С��ѷ��ţ�
				//�������������Ҫ�������飬�ñ����Ѿ��������д�������
				//������ѷ��ţ�����Ҫ��ʾ���ñ����ѷ��ţ��Ƿ����棿���û�ѡ���ǡ�����֮ǰ�������߼�����ѡ�񡰷��������
				//�����working״̬������Ҫ��ʾ���ñ����Ѵ��ڣ��Ƿ񸲸ǲ����������û�ѡ���ǡ���ֱ�Ӹ���MEDocument�汾�µ����ݼ���ѡ�񡰷��������
														
				//�ж�Ȩ��
//				boolean privilege = TCUtils.getTCSession().getTCAccessControlService().checkPrivilege(b8MEDocumentRev,"WRITE"); //$NON-NLS-1$
//				if(privilege){
					
					//�ж϶����Ƿ���������   �������������Ҫ�������飬�ñ����Ѿ��������д�������
					boolean fnd0InProcess = b8MEDocumentRev.getLogicalProperty("fnd0InProcess");
					
					if(fnd0InProcess){
						throw new TCException(Messages.ReportUtils_1); 
					}
									
					//�ж϶����Ƿ����״̬  ������ѷ��ţ�����Ҫ��ʾ���ñ����ѷ��ţ��Ƿ����棿���û�ѡ���ǡ�����֮ǰ�������߼�����ѡ�񡰷��������
					boolean state = TCComponentReleaseStatusUtils.existStatus(b8MEDocumentRev);
					
					if(state){
//						int num = JOptionPane.showConfirmDialog(null,
//								Messages.ReportUtils_2 + "��", //$NON-NLS-2$
//								Messages.ReportUtils_2,
//								JOptionPane.YES_NO_OPTION);						
//						
//						if (num == 0) {//��
							//����
							info.setAction("saveas");//����MEDocuemnt������Ҫ����
//						} else {//��
//							info.setIsgoon(false);
//						}
					} else {
//						int num = JOptionPane.showConfirmDialog(null,
//								Messages.ReportUtils_3 + "?", //$NON-NLS-2$
//								Messages.ReportUtils_3,
//								JOptionPane.YES_NO_OPTION);
//						if (num == 0) {//��
							//����
							info.setAction("replace");//����MEDocuemnt������Ҫ����
//						} else {//��
//							info.setIsgoon(false);
//						}
					}
//				} else {
//					throw new TCException(Messages.ReportUtils_4);
//				}
			} else {
				info.setAction("create");//�����ڵ�MEDocuemnt������Ҫ�´���
			}
		
		info.setMeDocument(b8MEDocumentRev);
		return info;
	}

	/* *************************
	 * ���������ɵı������TC
	 */
	public static TCComponentItem  afterGenerateReportAction(List<TCComponentDataset> datasetList, List<TCComponentItemRevision> productBOPList,
			GenerateReportInfo info,
			String documentName, String documentDesc, TCSession session) throws TCException {
		List<TCComponent> assignProjectComp = new ArrayList<TCComponent>();
		List<TCComponentProject> projects = new LinkedList<TCComponentProject>();
		boolean needToAssignProj = false;
		TCComponentItem dfl9MEDocumentItem = null;
		if(datasetList.size() > 0){
			String action = info.getAction();
			String dfl9_process_type = info.getDFL9_process_type();
			String dfl9_process_file_type = info.getDFL9_process_file_type();
			TCComponentItemRevision dfl9MEDocumentRev = info.getMeDocument();		
			if(dfl9MEDocumentRev!=null) {
				dfl9MEDocumentItem  = dfl9MEDocumentRev.getItem();
				//�����ĵ�����
				String newobjectname = info.getmeDocumentName();
				if(!newobjectname.isEmpty()) {
					dfl9MEDocumentItem.setProperty("object_name", newobjectname);
					dfl9MEDocumentItem.lock();
					dfl9MEDocumentItem.save();
					dfl9MEDocumentItem.unlock();
					dfl9MEDocumentRev.setProperty("object_name", newobjectname);
					dfl9MEDocumentRev.lock();
					dfl9MEDocumentRev.save();
					dfl9MEDocumentRev.unlock();
				}
			}else {
				dfl9MEDocumentItem = null;
			}
			
			Boolean isExist = info.isExist();
			
			if("create".equals(action)){	
				//������excel�ļ�����ΪMSExcelX���ݼ���ʹ�ù���ϵ���ص�һ���½�MEDocument�汾
				Map<String, Object> itemMap = new HashMap<String, Object>();
				Map<String, Object> itemRevisionMap = new HashMap<String, Object>();
				Map<String, Object> itemRevMasterFormMap = new HashMap<String, Object>();
				itemMap.put("item_id", ""); //$NON-NLS-1$ //$NON-NLS-2$
				itemMap.put("object_name", documentName); //$NON-NLS-1$
				itemMap.put("object_desc", documentDesc); //$NON-NLS-1$
				itemMap.put("object_type", DFL9MEDocument); //$NON-NLS-1$
				itemRevisionMap.put("object_type", DFL9MEDocumentRevision); //$NON-NLS-1$
				itemRevisionMap.put("object_name", documentName); //$NON-NLS-1$
				itemRevisionMap.put("dfl9_process_type", dfl9_process_type); //$NON-NLS-1$
				itemRevisionMap.put("dfl9_process_file_type", dfl9_process_file_type); //$NON-NLS-1$
				itemRevMasterFormMap.put("object_type", DFL9MEDocumentRevisionMaster); //$NON-NLS-1$
				
//				TCComponentItemType tcccomponentitemtype = (TCComponentItemType) session.getTypeComponent("DFL9MEDocument");
//
//				dfl9MEDocumentItem = tcccomponentitemtype.create("", "", "DFL9MEDocument", dfl9_process_file_type, "desc",
//						null);				
//				dfl9MEDocumentRev = dfl9MEDocumentItem.getLatestItemRevision();	
				
				try {
					CreateResponse respose = TCComponentUtils.create(itemMap, itemRevisionMap, itemRevMasterFormMap);
					int num = respose.serviceData.sizeOfCreatedObjects();
					if(num > 0){
						for(int i=0;i<num;i++){
							TCComponent comp = respose.serviceData.getCreatedObject(i);
							if(comp instanceof TCComponentItemRevision){
								dfl9MEDocumentRev = (TCComponentItemRevision) comp;						
							}
							
						}
					}
				} catch (ServiceException e) {
					e.printStackTrace();
					throw new TCException("Create " + ReportUtils.DFL9MEDocument +  " Fail : " +e.getMessage());  //$NON-NLS-1$ //$NON-NLS-2$
				}
				needToAssignProj = true;
				assignProjectComp.add(dfl9MEDocumentRev.getItem());
			}
			
			if("saveas".equals(action)){
				// ����MEDocument�汾���󣨲�����ԭ����ϵ��MSExcelX���ݼ���
				DeepCopyInfo deepCopyInfo = new DeepCopyInfo(dfl9MEDocumentRev, 1, null, null, false, false, false);
				deepCopyInfo.setAction(2);
				
				TCComponentItemRevision newRev = dfl9MEDocumentRev.saveAs("", dfl9MEDocumentRev.getStringProperty("object_name"),  dfl9MEDocumentRev.getStringProperty("object_desc"), false, new DeepCopyInfo[]{deepCopyInfo});//id name desc //$NON-NLS-1$ //$NON-NLS-2$ //$NON-NLS-3$
				// ��BOP�汾���Ƴ�ԭ�ϰ汾MEDocument�汾����ʹ�����ù�ϵ������MEDocument�汾����  
				for(TCComponentItemRevision prductBOPRev : productBOPList){
					//�Ƴ���ʱ����Ҫ�����з��������Ķ����ҳ��������Ƴ�  
					TCComponent[] children = TCComponentUtils.getCompsByRelation(prductBOPRev, ReportUtils.IMAN_reference);
					for(TCComponent child : children){
						if(child instanceof TCComponentItemRevision && DFL9MEDocumentRevision.equals(child.getType())){
							TCComponentItemRevision rev = (TCComponentItemRevision)child;
							String process_type = rev.getStringProperty("dfl9_process_type"); //$NON-NLS-1$
							String process_file_type = rev.getStringProperty("dfl9_process_file_type"); //$NON-NLS-1$
							String documentname = rev.getStringProperty("object_name");
							if(info.isFlag()) {
								if(dfl9_process_type.equals(process_type) && dfl9_process_file_type.equals(process_file_type)){
									prductBOPRev.remove(ReportUtils.IMAN_reference, rev);
								}
							}else {
								if(dfl9_process_type.equals(process_type) && dfl9_process_file_type.equals(process_file_type)&&info.getmeDocumentName().equals(documentname)){
									prductBOPRev.remove(ReportUtils.IMAN_reference, rev);
								}
							}
						}
					}
				}
				dfl9MEDocumentRev = newRev;
				needToAssignProj = true;
			}
			
			if("replace".equals(action)){
				//�Ƴ���ʱ����Ҫ�����з��������Ķ����ҳ��������Ƴ�  
				TCComponent[] children = TCComponentUtils.getCompsByRelation(dfl9MEDocumentRev, "IMAN_specification");
				for(TCComponent child : children){
					if(child instanceof TCComponentDataset){
						TCComponentDataset dataset = (TCComponentDataset)child;
						dfl9MEDocumentRev.cutOperation("IMAN_specification", new
								  TCComponent[]{dataset}); 						
						try{
							dataset.delete(); 
						}catch(Exception e2)
						{
							
						}
					}
				}
				needToAssignProj = true;
			}
			
			if(dfl9MEDocumentRev != null){
				dfl9MEDocumentRev.refresh();
				assignProjectComp.add(dfl9MEDocumentRev);
				// ������excel�ļ�����ΪMSExcelX���ݼ���ʹ�ù���ϵ���ص������İ汾�����£�
				for(TCComponentDataset dataset : datasetList){
					TCComponentUtils.createRelation(dfl9MEDocumentRev, dataset, "IMAN_specification");
				}
				
				//����Ӧ��MEDocument�汾����ʹ�����ù�ϵ���ص����е�ProductBOP�汾�£�
				for(TCComponentItemRevision prductBOPRev : productBOPList){
					if(info.getProject_ids()==null) {
						TCComponent[] comps = TCComponentProjectUtils.getTCComponentProjects(prductBOPRev);
						if(comps != null && comps.length > 0){
							for(TCComponent comp : comps){
								if(comp instanceof TCComponentProject){
									projects.add((TCComponentProject)comp);
								}
							}
						}	
					}else {
						TCComponent[] comps = TCComponentProjectUtils.getTCComponentProjects(info.getProject_ids());
						if(comps != null && comps.length > 0){
							for(TCComponent comp : comps){
								if(comp instanceof TCComponentProject){
									projects.add((TCComponentProject)comp);
								}
							}
						}	
					}
					//��item���ص�BOP�汾��
					dfl9MEDocumentItem = dfl9MEDocumentRev.getItem();
					TCComponentUtils.createRelation(prductBOPRev, dfl9MEDocumentItem, ReportUtils.IMAN_reference);
				}
			} else {
				if(isExist){
					throw new TCException("SaveAs " + ReportUtils.DFL9MEDocument +  " Fail");  //$NON-NLS-1$ //$NON-NLS-2$
				} else {
					throw new TCException("Create " + ReportUtils.DFL9MEDocument +  " Fail");  //$NON-NLS-1$ //$NON-NLS-2$
				}
			}
			
			if(needToAssignProj && assignProjectComp.size() > 0 && projects.size() > 0){
				try {
					for(TCComponentProject proj : projects){
						TCComponentProjectUtils.assignProject(assignProjectComp, proj);
					}
				} catch (TCException e) {
					e.printStackTrace();
					System.out.println("assignProject TCException : " + e.getMessage()); //$NON-NLS-1$
				}
			}
			
		}
		return dfl9MEDocumentItem;
	}	

	/* *************************
	 * ����BOP���ƻ�ȡ��������
	 */
	public static String getFactoryLineByBOP(String str) {
		String factoryline = "";
		String[] values = str.split("_");
		if(values.length>3) {
			factoryline = values[2];
		}
		return factoryline;		
	}
	
	/**
	 * /**
	 * Excel2007
	 * ��ȡָ�������ڵ�ͼƬ����
	 * @param sheet	��ǰsheet���
	 * @param wb	����������
	 * @param beginRow	����������ʼ��
	 * @param endRow	 ����������ֹ��
	 * @param beginCol	����������ʼ��
	 * @param endCol	 ����������ֹ��
	 * @return List<String> ͼƬ�����б�
	 */
	public static List<String> removePictrues07(XSSFSheet sheet,
			XSSFWorkbook wb, int beginRow, int endRow, int beginCol, int endCol) {
		System.out.println();
		System.out.println("ָ������" 
				+ String.valueOf(beginRow) + ","
				+ String.valueOf(endRow)+ ","
				+ String.valueOf(beginCol) + ","
				+ String.valueOf(endCol));
		List<String> delPicturesList = new ArrayList<String>();
		List<POIXMLDocumentPart> relations = sheet.getRelations();
		for (int i = 0; i < relations.size(); i++) {
			POIXMLDocumentPart dr = relations.get(i);
			if (dr instanceof XSSFDrawing) {
				XSSFDrawing drawing = (XSSFDrawing) dr;
				List<XSSFShape> shapes = drawing.getShapes();
				for (XSSFShape shape : shapes) {
					if (shape instanceof XSSFPicture) {
						XSSFPicture pic = (XSSFPicture) shape;
						if(pic.getAnchor()!=null) {
							XSSFClientAnchor anchor = pic.getPreferredSize();						
							System.out.println("target picture��" 
									+ String.valueOf(anchor.getRow1()) + ","
									+ String.valueOf(anchor.getRow2())+ ","
									+ String.valueOf(anchor.getCol1()) + ","
									+ String.valueOf(anchor.getCol2()));
							String name = pic.getCTPicture().getNvPicPr()
									.getCNvPr().getName();
							if (isCellInScope(anchor.getRow1(), anchor.getCol1(),
									beginRow, endRow, beginCol, endCol)) {							
								delPicturesList.add(name);
								System.out.println(name + " ��ָ��������");
							} else {
								System.out.println(name + " ����ָ��������");
							}
						}					
					}
				}

			}
		}
		System.out.println();
		return delPicturesList;
	}
	
	/**
	 * �ж�Ŀ��ͼƬ��ʼ��Ԫ��λ���Ƿ��ڸ�������Χ��
	 * 
	 * @param tagrgetRow
	 *            Ŀ��ͼƬ������
	 * @param tagrgetCol
	 *            Ŀ��ͼƬ������
	 * @param row
	 *            ����������ʼ��
	 * @param toRow
	 *            ����������ֹ��
	 * @param col
	 *            ����������ʼ��
	 * @param toCol
	 *            ����������ֹ��
	 * @return
	 */
	public static boolean isCellInScope(int tagrgetRow, int tagrgetCol,
			int row, int toRow, int col, int toCol) {
		if (row <= tagrgetRow && tagrgetRow <= toRow) {
			if (col <= tagrgetCol && tagrgetCol <= toCol) {
				return true;
			}
		}
		return false;
	}
	

	public static String getCellValueString(Cell cell) {
		if (cell == null) {
			return null;
		}
		String returnValue = "";
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC: // ����
			//			cell.getStringCellValue();
			Double doubleValue = cell.getNumericCellValue();
			// ��ʽ����ѧ��������ȡһλ����
			DecimalFormat df = new DecimalFormat("0");
			returnValue = df.format(doubleValue);
			break;
		case Cell.CELL_TYPE_STRING: // �ַ���
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
		return returnValue.trim();
	}
	
	
	public static Workbook getWorkbook(String excelPath) {

		File excelFile = new File(excelPath);
		if (!excelFile.exists()) {
			System.out.println("excel�����ڣ�" + excelPath);
			return null;
		}

		String fileName = excelFile.getName();
		Workbook workbook = null;

		try {
			FileInputStream inputStream = new FileInputStream(excelFile);
			if (fileName.endsWith("xls")) {

				workbook = new HSSFWorkbook(inputStream);

			} else if (fileName.endsWith("xlsx")) {
				workbook = new XSSFWorkbook(inputStream);
			}

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			System.out.println(e.getMessage());
			System.out.println(e);
		}

		return workbook;
	}
	
}
