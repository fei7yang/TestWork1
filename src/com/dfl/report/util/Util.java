package com.dfl.report.util;

import java.awt.Color;
import java.awt.Graphics;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.imageio.ImageIO;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;
import org.eclipse.core.runtime.Plugin;
import org.eclipse.swt.graphics.Image;

import com.dfl.report.Activator.Activator;
import com.dfl9.services.loose.weldasm.SetByPassService;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.designcontext.util.ISearchParameter;
import com.teamcenter.rac.designcontext.util.ItemAttributeSearchParameter;
import com.teamcenter.rac.kernel.ServiceData;
import com.teamcenter.rac.kernel.SoaUtil;
import com.teamcenter.rac.kernel.TCAccessControlService;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentDatasetType;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentForm;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentItemType;
import com.teamcenter.rac.kernel.TCComponentProject;
import com.teamcenter.rac.kernel.TCComponentQuery;
import com.teamcenter.rac.kernel.TCComponentQueryType;
import com.teamcenter.rac.kernel.TCComponentTcFile;
import com.teamcenter.rac.kernel.TCComponentUser;
import com.teamcenter.rac.kernel.TCDateFormat;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCProperty;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.kernel.TCUserService;
import com.teamcenter.rac.pse.common.BOMLineNode;
import com.teamcenter.rac.psebase.AbstractBOMLineViewerApplication;
import com.teamcenter.rac.psebase.common.AbstractViewableTreeTable;
import com.teamcenter.rac.testmanager.Messages;
import com.teamcenter.rac.testmanager.TSMUtils;
import com.teamcenter.rac.testmanager.steps.CheckPlacemarksFilledOperation;
import com.teamcenter.rac.testmanager.ui.model.TestManagerDataCache;
import com.teamcenter.rac.testmanager.ui.model.TestManagerModelObject;
import com.teamcenter.rac.util.ConfirmationDialog;
import com.teamcenter.rac.util.OSGIUtil;
import com.teamcenter.services.rac.structuremanagement.StructureFilterWithExpandService;
import com.teamcenter.soa.exceptions.NotLoadedException;
import com.tm0.services.internal.rac.testmanagement.InstanceManagementService;

import com.teamcenter.rac.structure.search.model.StructureSearchCriteriaModel;
import com.teamcenter.rac.structure.search.model.StructureSearchModel;
import com.teamcenter.rac.structure.search.model.StructureSearchResultsModel;
import com.teamcenter.rac.structure.search.services.IStructureSearchService;
import com.teamcenter.services.rac.core.DataManagementService;
import com.teamcenter.services.rac.core._2007_01.DataManagement;
import com.teamcenter.services.rac.manufacturing.StructureSearchService;
import com.teamcenter.services.rac.structuremanagement.StructureFilterWithExpandService;

public class Util {
	
	
	public static String getPropertyDisplayValue(TCComponent component, String property)
	{
		String value = "";
		try {
			value =component.getPropertyDisplayableValue(property);
		} catch (NotLoadedException e) {
			// TODO Auto-generated catch block
//			e.printStackTrace();
			return getProperty(component, property);
		}
		return value ;
	}
	/**
	 * 获取关系对象数组
	 * 
	 * @param comp
	 * @param relation
	 * @return
	 */
	public static TCComponent[] getRelComponents(TCComponent comp, String relation) {
		TCComponent[] comps = null;
		try {
			if (comp == null) {
				return comps;
			}
			comps = comp.getRelatedComponents(relation);
		} catch (TCException e) {
			// TODO Auto-generated catch block
			System.out.println("ERROR:" + comp.toString());
			e.printStackTrace();
		}

		return comps;
	}

	/**
	 * 获取关系对象
	 * 
	 * @param comp
	 * @param relation
	 * @return
	 */
	public static TCComponent getRelComponent(TCComponent comp, String relation) {
		TCComponent component = null;
		try {

			if (comp == null) {
				return component;
			}

			TCComponent[] comps = comp.getRelatedComponents(relation);
			if (comps != null && comps.length > 0) {
				component = comps[0];// 如果存在多个制造目标怎么解决 ？？？
			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			System.out.println("ERROR:" + comp.toString());
			e.printStackTrace();
		}

		return component;
	}

	/**
	 * 获取对象的属性
	 * 
	 * @param target2
	 * @param string
	 * @return
	 */
	public static String getProperty(TCComponent component, String property) {
		if (component == null) {
			return "";
		}

//		try {
//			component.refresh();
//		} catch (TCException e) {
//		}

		try {
			TCProperty p = component.getTCProperty(property);
			if (p == null) {
				return "";
			}
			if (property.equals("")) {
				return "";
			}
			String s = component.getProperty(property);
			if (s == null) {
				return "";
			} else {
				return s;
			}

		} catch (TCException e) {
			e.printStackTrace();
		}
		return "";
	}

	/**
	 * 获取对象的属性
	 * 
	 * @param target2
	 * @param string
	 * @return
	 */
	public static String getRelProperty(TCComponent component, String property) {
		if (component == null) {
			return "";
		}

//		try {
//			component.refresh();
//		} catch (TCException e) {
//		}

		try {
			TCProperty p = component.getTCProperty(property);
			if (p == null) {
				return "";
			}
			if (property.equals("")) {
				return "";
			}
			// String s = component.getProperty(property);
			String s = p.getStringValue();
			String s2 = p.getDisplayValue();
			System.out.println("S=" + s + " S2=" + s2);
			if (s == null) {
				return "";
			} else {
				return s;
			}

		} catch (TCException e) {
			e.printStackTrace();
		}
		return "";
	}

	/**
	 * 检查对指定对象的写权限
	 * 
	 * @param session
	 * @param comp
	 * @return
	 */
	public static boolean hasWritePrivilege(TCSession session, TCComponent comp) {
		try {
			TCAccessControlService accessControlService = session.getTCAccessControlService();
			TCComponent component = (TCComponent) comp;

			if (comp instanceof TCComponentBOMLine) {
				// 检查bomview
				TCComponentBOMLine bomline = (TCComponentBOMLine) comp;
				TCComponentItemRevision ir = bomline.getItemRevision();
				if (accessControlService.checkPrivilege(ir, "WRITE")) {
					TCComponent[] bvs = ir.getRelatedComponents("structure_revisions");
					if (bvs != null && bvs.length > 0) {
						component = bvs[0];
					}
				} else {
					return false;
				}
				// component =
			}

			return accessControlService.checkPrivilege(component, "WRITE");
		} catch (TCException e) {
			e.printStackTrace();
		}
		return false;
	}

	/**
	 * 获取布尔类型的属性值
	 * 
	 * @param component
	 * @param property
	 * @return
	 */
	public static boolean getLogicalProperty(TCComponent component, String property) {
		// TODO Auto-generated method stub
		if (component == null) {
			return false;
		}
		try {
			boolean flag = component.getLogicalProperty(property);
			return flag;
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return false;
	}

	/**
	 * 判断对象是否已有状态
	 * 
	 * @param object
	 * @return
	 */
	public static boolean hasStatus(TCComponent object) {
		try {

			object.refresh();
			TCComponent[] relStatus = object.getReferenceListProperty("release_status_list");
			if (relStatus == null || relStatus.length <= 0) {
				// System.out.println("找不到状态");
				return false;
			} else {
				return true;
			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			return false;
		}

	}

	public static TCComponentDataset createDatasetKeepFile(TCSession session, String name, String filePath, String dsType,
			String nameRef) {
		// TODO Auto-generated method stub
		TCComponentDataset dataset = null;
		try {

			TCComponentDatasetType datasetType = (TCComponentDatasetType) session.getTypeComponent("Dataset");
			dataset = datasetType.create(name, "", dsType);

			String[] filePaths = { filePath };
			String[] namedRefs = { nameRef };
			dataset.setFiles(filePaths, namedRefs);
			dataset.lock();
			dataset.save();
			dataset.unlock();

			// 删除中间文件
//			File file = new File(filePath);
//			if (file.isFile()) {
//				file.delete();
//			}

		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			return null;
		}
		return dataset;
	}
	
	public static TCComponentDataset createDataset(TCSession session, String name, String filePath, String dsType,
			String nameRef) {
		// TODO Auto-generated method stub
		TCComponentDataset dataset = null;
		try {

			TCComponentDatasetType datasetType = (TCComponentDatasetType) session.getTypeComponent("Dataset");
			dataset = datasetType.create(name, "", dsType);

			String[] filePaths = { filePath };
			String[] namedRefs = { nameRef };
			dataset.setFiles(filePaths, namedRefs);
			dataset.lock();
			dataset.save();
			dataset.unlock();

			// 删除中间文件
			File file = new File(filePath);
			if (file.isFile()) {
				file.delete();
			}

		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			return null;
		}
		return dataset;
	}

	/* 搜索获取测试用例和测试用例实例 */
	public static ArrayList SearchTests(TCComponentBOMLine m_srchScope) throws Exception {

		ArrayList resultlist = new ArrayList();
		String m_resultStatus = "";
		// String m_testType =
		// TestManagerDataCache.getRealValueForDisplayValue("Tm0TestType", m_testType);
		// String m_resultStatus =
		// TestManagerDataCache.getRealValueForDisplayValue("Tm0ResultStatus",
		// m_resultStatus);

		com.tm0.services.internal.rac.testmanagement._2015_03.InstanceManagement.TestInstancesSearchCriteria testinstancessearchcriteria = new com.tm0.services.internal.rac.testmanagement._2015_03.InstanceManagement.TestInstancesSearchCriteria();
		testinstancessearchcriteria.clientId = "searchTests";
		testinstancessearchcriteria.searchScope = m_srchScope;
		testinstancessearchcriteria.executionScope = m_srchScope;// m_execScope
		testinstancessearchcriteria.testCase = null;// m_testCase
		testinstancessearchcriteria.referredObjects = new TCComponent[0];// m_refObjs

		InstanceManagementService instancemanagementservice = InstanceManagementService.getService(TSMUtils.TCSESSION);

		if (instancemanagementservice == null) {
			System.out.println("instancemanagementservice==null");
			return resultlist;
		}

		Object obj = null;
		com.tm0.services.internal.rac.testmanagement._2015_03.InstanceManagement.TestInstancesSearchResponse testinstancessearchresponse;

		// 搜索测试
		try {
			TSMUtils.setPropertyPolicy(String.format("%s,%s,%s,%s,%s",
					new Object[] { "OBJ_PROP_POLICY_TESTINSTANCE", "OBJ_PROP_POLICY_TESTACTIVITY",
							"OBJ_PROP_POLICY_BOMLINE", "OBJ_PROP_POLICY_ITEM", "OBJ_PROP_POLICY_ITEMREV" }));
			com.tm0.services.internal.rac.testmanagement._2015_03.InstanceManagement.TestInstancesSearchCriteria[] instancesSearchCriterias = new com.tm0.services.internal.rac.testmanagement._2015_03.InstanceManagement.TestInstancesSearchCriteria[1];
			instancesSearchCriterias[0] = testinstancessearchcriteria;
			testinstancessearchresponse = instancemanagementservice.getTestInstances(instancesSearchCriterias);
		} finally {
			TSMUtils.restorePropertyPolicy();
		}

		// 获取搜索结果
		SoaUtil.handlePartialErrors(testinstancessearchresponse.srvcData, null);
		ArrayList arraylist = new ArrayList();
		TCDateFormat tcdateformat = new TCDateFormat(TSMUtils.TCSESSION);
		Object obj1 = testinstancessearchresponse.resultMap.get(testinstancessearchcriteria.clientId);
		if (!(obj1 instanceof com.tm0.services.internal.rac.testmanagement._2015_03.InstanceManagement.TestInstancesSearchResult[])) {
			return null;
		}
		HashMap hashmap = new HashMap();
		com.tm0.services.internal.rac.testmanagement._2015_03.InstanceManagement.TestInstancesSearchResult atestinstancessearchresult[] = (com.tm0.services.internal.rac.testmanagement._2015_03.InstanceManagement.TestInstancesSearchResult[]) obj1;
		com.tm0.services.internal.rac.testmanagement._2015_03.InstanceManagement.TestInstancesSearchResult atestinstancessearchresult1[];
		int j = (atestinstancessearchresult1 = atestinstancessearchresult).length;

		System.out.println("本次搜索结果数量：" + j);

		for (int i = 0; i < j; i++) {
			com.tm0.services.internal.rac.testmanagement._2015_03.InstanceManagement.TestInstancesSearchResult testinstancessearchresult = atestinstancessearchresult1[i];
			TCComponent tccomponent1 = testinstancessearchresult.testInstance;
			TCComponent tccomponent2 = testinstancessearchresult.executionScope;
			if (tccomponent1 == null || tccomponent2 == null) {
				System.out.println("Test Instance or Exec Scope is NULL in the search result. Ignoring");
				continue;
			}
			System.out.println("testInstance:" + tccomponent1.toString());

			System.out.println("testInstance type:" + tccomponent1.getType());

			String s1 = "";
			TCProperty tcproperty = null;
			if (testinstancessearchresult.activities.length > 0) {
				TCComponent tccomponent3 = testinstancessearchresult.activities[0];
				System.out.println("activities:" + tccomponent3.toString());
				System.out.println("activities type:" + tccomponent3.getType());

				if (tccomponent3 != null) {
					Date date = tccomponent3.getDateProperty("tm0ActivityDate");
					if (date != null) {
						// if(m_searchDate != null && m_searchDate.getTime().compareTo(date) < 0)
						// continue;
						s1 = tcdateformat.format(date);
					}
					tcproperty = tccomponent3.getTCProperty("tm0ResultStatus");
					if (tcproperty == null || tcproperty.getStringValue().isEmpty()) {
						tcproperty = tccomponent3.getTCProperty("tm0ActivityType");
					}
				}

			}

			String s3 = Messages.TSMView_noResultStatus_TEXT;
			if (tcproperty != null && !tcproperty.getStringValue().isEmpty()) {
				s3 = tcproperty.getDisplayableValue();
			}
			TestManagerModelObject testmanagermodelobject1 = new TestManagerModelObject(tccomponent1, tccomponent2,
					testinstancessearchresult.assignedTo, s1, s3);
			hashmap.put(tccomponent1, testmanagermodelobject1);
			com.tm0.services.internal.rac.testmanagement._2015_10.InstanceManagement.TestStepsInputInfo teststepsinputinfo = new com.tm0.services.internal.rac.testmanagement._2015_10.InstanceManagement.TestStepsInputInfo();
			teststepsinputinfo.testObject = tccomponent1;
			teststepsinputinfo.bomLineScope = tccomponent2;
			arraylist.add(teststepsinputinfo);
		}

		CheckPlacemarksFilledOperation checkplacemarksfilledoperation = new CheckPlacemarksFilledOperation(arraylist);
		checkplacemarksfilledoperation.executeOperation();
		Iterator iterator = hashmap.entrySet().iterator();
		while (iterator.hasNext()) {
			java.util.Map.Entry entry = (java.util.Map.Entry) iterator.next();
			TCComponent tccomponent = (TCComponent) entry.getKey();
			TestManagerModelObject testmanagermodelobject = (TestManagerModelObject) entry.getValue();
			String s = checkplacemarksfilledoperation.getPlaceHolderStatus(tccomponent);
			if ("Requires Detailing".equals(s) || "Missing Object Reference".equals(s)) {
				if (m_resultStatus != null && !m_resultStatus.isEmpty()
						&& !m_resultStatus.equals(Messages.TSMView_noResultStatus_TEXT)) {
					continue;
				}
				testmanagermodelobject.setOverrideStatus(s);
			} else if (m_resultStatus != null && !m_resultStatus.isEmpty()) {
				String s2 = TestManagerDataCache.getRealValueForDisplayValue("Tm0ResultStatus",
						testmanagermodelobject.getResultStatus());
				if (!m_resultStatus.equals(s2)) {
					continue;
				}
			}
			resultlist.add(testmanagermodelobject);
		}
		hashmap.clear();

		return resultlist;
	}

	/**
	 * 搜索BOM行 ，获取满足条件的bomline
	 */
	public static ArrayList searchBOMLine(TCComponentBOMLine parent, String logicalOperator, String[] propertys,
			String operator, String values[]) {
		// TODO Auto-generated method stub

		com.teamcenter.services.rac.structuremanagement._2014_06.StructureFilterWithExpand.ExpandAndSearchResponse expandandsearchresponse = null;
		StructureFilterWithExpandService structurefilterwithexpandservice = StructureFilterWithExpandService
				.getService(parent.getSession());

		int l = values.length;
		com.teamcenter.services.rac.structuremanagement._2014_06.StructureFilterWithExpand.SearchCondition[] conditions = new com.teamcenter.services.rac.structuremanagement._2014_06.StructureFilterWithExpand.SearchCondition[l];
		for (int i = 0; i < l; i++) {
			conditions[i] = new com.teamcenter.services.rac.structuremanagement._2014_06.StructureFilterWithExpand.SearchCondition();
			// if(i>0){
			conditions[i].logicalOperator = logicalOperator;
			// }else{
			// conditions[i].logicalOperator = "";
			// }
			conditions[i].propertyName = propertys[i];
			conditions[i].relationalOperator = operator;
			conditions[i].inputValue = values[i];
		}

		TCComponentBOMLine atccomponentbomline[] = { parent };
		expandandsearchresponse = structurefilterwithexpandservice.expandAndSearch(atccomponentbomline, conditions);

		ArrayList arraylist = new ArrayList();

		for (int i = 0; i < expandandsearchresponse.outputLines.length; i++) {
			TCComponentBOMLine line = expandandsearchresponse.outputLines[i].resultLine;
			try {
				String type = line.getItem().getType();
				/*
				 * if(type.equals("MEWorkarea")||type.equals("MEProductLocation" )) { continue;
				 * } if(line.parent()!=null) { TCComponentBOMLine parentLine = line.parent();
				 * String parentType = parentLine.getItem().getType();
				 * if(parentType.equals("WH3_SBOMPart")) { continue; } }
				 */
				arraylist.add(expandandsearchresponse.outputLines[i].resultLine);

			} catch (TCException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

		}

		return arraylist;

	}

	/**
	 * 获取用户名称
	 */

	public static String getUserName(TCComponent comp) {
		try {
			TCComponentUser user = (TCComponentUser) comp.getReferenceProperty("owning_user");

			return getProperty(user, "user_name");

		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return "";
	}

	/**
	 * 通过BONLine获取下一级的子集
	 * 
	 * @param session
	 * @param comp
	 * @return
	 */
	public static ArrayList getChildrenByParent(TCComponentBOMLine parent) {
		ArrayList list = new ArrayList();
		try {
			AIFComponentContext[] childrens = parent.getChildren();
			for (AIFComponentContext chil : childrens) {
				TCComponentBOMLine bl = (TCComponentBOMLine) chil.getComponent();
				list.add(bl);
			}
			return list;

		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return list;
	}

	/**
	 * 通过BONLine根据制定的类型返回其下一级的子集
	 * 
	 * @param session
	 * @param comp
	 * @return
	 */
	public static ArrayList<TCComponentBOMLine> getChildrenByBOMLine(TCComponentBOMLine parent, String object_type) {
		ArrayList list = new ArrayList();
		try {
			AIFComponentContext[] childrens = parent.getChildren();
			for (AIFComponentContext chil : childrens) {
				TCComponentBOMLine bl = (TCComponentBOMLine) chil.getComponent();
				TCComponentItemRevision rev = (TCComponentItemRevision) bl.getItemRevision();
				if (rev != null) {
					if (rev.isTypeOf(object_type)) {
						list.add(bl);
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

	// 多张图片合成一张
	public static void batchmergeImage(String[] filestr, String fileStr3, String path, int totalwidth, int totalheight)
			throws IOException {
		if (filestr != null && filestr.length > 0) {

			// 根据图片的数量来排版
			int rows = (filestr.length + 1) / 2;// 排多少行
			// 取第一张图片的宽和高位标准
			int basewidth = totalwidth / 2;
			int basehight = totalheight / rows;

			BufferedImage combined = new BufferedImage(totalwidth, totalheight, BufferedImage.TYPE_INT_RGB);

			Graphics g = combined.getGraphics();
			try {
				for (int i = 0; i < filestr.length; i++) {
					File file = new File(path, filestr[i]);
					BufferedImage image = ImgCompress(file, basewidth, basehight);
					if ((i + 1) % 2 == 0) {
						g.drawImage(image, basewidth, basehight * (i - 1) / 2, null);
						System.out.println(basewidth + " " + basehight * (i - 1) / 2);
					} else {
						g.drawImage(image, 0, basehight * i / 2, null);
						System.out.println("0" + " " + basehight * i / 2);
					}
				}
				// 如果图片数量为单数，需要补白
				if (filestr.length % 2 != 0) {

					g.setColor(Color.white);
					g.fillRect(basewidth, basehight * (rows - 1), basewidth, basehight);
					g.dispose();
				}
				// Save as new image
				ImageIO.write(combined, "png", new File(path, fileStr3));
			} finally {
				if (g != null) {
					g.dispose();
				}
			}

		}
	}

	// 图片压缩处理

	private static BufferedImage img;
	private static int width;
	private static int height;

	public static BufferedImage ImgCompress(File file, int basewidth, int baseheight) throws IOException {
		img = ImageIO.read(file); // 构造Image对象
		width = img.getWidth(null); // 得到源图宽
		height = img.getHeight(null); // 得到源图长

//		    return resizeFix(400, 492);
		return resize(basewidth, baseheight);
	}

	/**
	 * 按照宽度还是高度进行压缩
	 * 
	 * @param w int 最大宽度
	 * @param h int 最大高度
	 */
	public static BufferedImage resizeFix(int w, int h) throws IOException {
		if (width / height > w / h) {
			return resizeByWidth(w);
		} else {
			return resizeByHeight(h);
		}
	}

	/**
	 * 以宽度为基准，等比例放缩图片
	 * 
	 * @param w int 新宽度
	 */
	public static BufferedImage resizeByWidth(int w) throws IOException {
		int h = (int) (height * w / width);
		return resize(w, h);
	}

	/**
	 * 以高度为基准，等比例缩放图片
	 * 
	 * @param h int 新高度
	 */
	public static BufferedImage resizeByHeight(int h) throws IOException {
		int w = (int) (width * h / height);
		return resize(w, h);
	}

	/**
	 * 强制压缩/放大图片到固定的大小
	 * 
	 * @param w int 新宽度
	 * @param h int 新高度
	 */
	public static BufferedImage resize(int w, int h) throws IOException {
		// SCALE_SMOOTH 的缩略算法 生成缩略图片的平滑度的 优先级比速度高 生成的图片质量比较好 但速度慢
		BufferedImage image = new BufferedImage(w, h, BufferedImage.TYPE_INT_RGB);
		Graphics g = image.getGraphics();
		try {
			g.drawImage(img, 0, 0, w, h, null); // 绘制缩小后的图
		} finally {
			if (g != null) {
				g.dispose();
			}
		}
		return image;
	}

	/**
	 * 获取bodytext的显示文本
	 * 
	 * @param bodyText xml格式的字符串
	 * @return
	 */
	public static String getBodyText(String bodyText) {
		String displayValue = "";

		if (bodyText.trim().equals("")) {
			return displayValue;
		}

		Document doc = initDoc(bodyText);
		if (doc == null) {
			return displayValue;
		}

		Element root = doc.getRootElement();
		List<Element> elementList = root.elements();
		for (Element element : elementList) {
			// System.out.println("element:"+element.getName());
			if (element.getName().equals("Step")) {
				String value = element.elementText("TextSegment");
//	     System.out.println(value);
				if (displayValue.equals("")) {
					displayValue = displayValue + value;
				} else {
					displayValue = displayValue + "\n" + value;
				}
			}
		}

		return displayValue;
	}

	public static Document initDoc(String xmlContent) {
		Document doc = null;
		try {
			doc = (Document) DocumentHelper.parseText(xmlContent);
		} catch (DocumentException e) {
			// e.printStackTrace();
			System.out.println("XML解析错误：" + xmlContent);
		}
		return doc;
	}

	/**
	 * 下载图片数据集到本地
	 * 
	 * @param picDs1
	 * @return
	 */
	public static File downLoadPicture(TCComponent comp) {
		// TODO Auto-generated method stub

		// System.out.println(">>>downLoadPicture");

		TCComponentDataset dataset = null;
		if (comp instanceof TCComponentDataset) {
			dataset = (TCComponentDataset) comp;
		}
		File file = null;
		if (dataset == null) {
			// System.out.println("dataset==null");
			return null;
		}

		System.out.println("downLoadPicture:" + dataset.toString());
		String type = dataset.getType();
		// "Image","JPEG","Bitmap","TIF","GIF"
		if (!"Vis_Snapshot_2D_View_Data".equals(type) && !"SnapShotViewData".equals(type) && !"Image".equals(type)
				&& !"JPEG".equals(type) && !"Bitmap".equals(type) && !"TIF".equals(type) && !"GIF".equals(type)) {
			// System.out.println("图片类型不匹配："+type);
			return null;
		}

		TCComponentTcFile[] files;
		try {

			files = dataset.getTcFiles();
			if (files == null || files.length <= 0) {
				return null;
			}
			for (int i = 0; i < files.length; i++) {
				String fileName = files[i].getProperty("file_name");
				System.out.println("fileName:" + fileName);
				if (fileName.toLowerCase().endsWith("png") || fileName.toLowerCase().endsWith("jpeg")
						|| fileName.toLowerCase().endsWith("jpg") || fileName.toLowerCase().endsWith("bmp")
						|| fileName.toLowerCase().endsWith("tif") || fileName.toLowerCase().endsWith("gif")) {
					file = files[i].getFmsFile();
					// System.out.println("fms file:"+file.getAbsolutePath());
					return file;
				}
			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return file;
	}

	/*
	 * ****************************************************** 把生成的报表保存在指定的文件夹下
	 */
	public static void saveFiles(String filename, String datasetName, TCComponent folder, TCSession session,
			String type) {
		try {
			String fullFileName = FileUtil.getReportFileName(filename);
			TCComponentFolder savefolder = (TCComponentFolder) folder;

			TCComponentItemType tcccomponentitemtype = (TCComponentItemType) session.getTypeComponent("B8_BIWProcDoc");
			TCComponentItem tccomponentitem = tcccomponentitemtype.create("", "", "B8_BIWProcDoc", datasetName, "desc",
					null);
			tccomponentitem.setProperty("b8_BIWProcDocType", type);
			tccomponentitem.lock();
			tccomponentitem.save();
			tccomponentitem.unlock();
			TCComponentDataset ds = Util.createDataset(session, datasetName, fullFileName, "MSExcelX", "excel");
			// 添加文档与数据集的关系
			TCComponentItemRevision rev = tccomponentitem.getLatestItemRevision();
			rev.add("IMAN_specification", ds);
			rev.lock();
			rev.save();
			rev.unlock();
			savefolder.add("contents", tccomponentitem);
			// 删除中间文件
			File file = new File(fullFileName);
			if (file.isFile()) {
				file.delete();
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static List callStructureSearch(TCComponentBOMLine scopesBomLine, String queryName, String[] entries,
			String[] values) throws TCException {
		List list = new ArrayList();

		IStructureSearchService iStructureSearchService = (IStructureSearchService) OSGIUtil
				.getService((Plugin) Activator.getDefault(), IStructureSearchService.class);
		StructureSearchCriteriaModel structureSearchCriteriaModel = new StructureSearchCriteriaModel(new ArrayList(),
				new ArrayList());
		ItemAttributeSearchParameter itemParameter = null;
		ArrayList<ISearchParameter> searchParameterList = new ArrayList<ISearchParameter>();

		// 搜索范围(BOM顶层)
		{
			ArrayList scopeList = new ArrayList();
			scopeList.add(scopesBomLine);
			structureSearchCriteriaModel.setScopes(scopeList);
		}

		// 搜索条件
		{
			// Item属性
			itemParameter = new ItemAttributeSearchParameter();
			TCComponentQueryType queryType = (TCComponentQueryType) scopesBomLine.getSession()
					.getTypeComponent("ImanQuery");
			TCComponentQuery query = (TCComponentQuery) queryType.find(queryName);
			if (query == null) {
				System.out.println("ERROR:查询未定义：" + queryName);
				throw new TCException("ERROR：查询未定义");
				// return list;
			}
			// String[] entries = new String[]{"ID"};
			// String[] values = new String[]{"*"};
			itemParameter.setEntries(entries);
			itemParameter.setValues(values);
			itemParameter.setSavedQuery(query);
			searchParameterList.add((ISearchParameter) itemParameter);
		}

		structureSearchCriteriaModel.addParameters(searchParameterList);
		// 执行搜索
		StructureSearchModel structureSearchModel = iStructureSearchService
				.performSearch(structureSearchCriteriaModel.clone(), true);

		// 获取搜索结果
		StructureSearchResultsModel structureSearchResultsModel = structureSearchModel.getResults();
		if (structureSearchModel != null && structureSearchResultsModel != null) {
			list = structureSearchResultsModel.getResults();
			System.out.println("结构搜索结果数量:" + list.size());
		}

		return list;
	}

	public static List<TCComponent> callStructureSearch(List<TCComponent> scopeList, String queryName, String[] entries,
			String[] values) throws TCException {
		List<TCComponent> list = new ArrayList<TCComponent>();

		IStructureSearchService iStructureSearchService = (IStructureSearchService) OSGIUtil
				.getService((Plugin) Activator.getDefault(), IStructureSearchService.class);
		StructureSearchCriteriaModel structureSearchCriteriaModel = new StructureSearchCriteriaModel(new ArrayList(),
				new ArrayList());
		ItemAttributeSearchParameter itemParameter = null;
		ArrayList<ISearchParameter> searchParameterList = new ArrayList<ISearchParameter>();

		// 搜索条件
		{
			// Item属性
			itemParameter = new ItemAttributeSearchParameter();
			TCComponentQueryType queryType = (TCComponentQueryType) scopeList.get(0).getSession()
					.getTypeComponent("ImanQuery");
			TCComponentQuery query = (TCComponentQuery) queryType.find(queryName);
			if (query == null) {
				System.out.println("ERROR:查询未定义：" + queryName);
				throw new TCException("ERROR：查询未定义");
				// return list;
			}
			// String[] entries = new String[]{"ID"};
			// String[] values = new String[]{"*"};
			itemParameter.setEntries(entries);
			itemParameter.setValues(values);
			itemParameter.setSavedQuery(query);
			searchParameterList.add((ISearchParameter) itemParameter);
		}
		structureSearchCriteriaModel.setScopes(scopeList);
		structureSearchCriteriaModel.addParameters(searchParameterList);
		// 执行搜索
		StructureSearchModel structureSearchModel = iStructureSearchService
				.performSearch(structureSearchCriteriaModel.clone(), true);

		// 获取搜索结果
		StructureSearchResultsModel structureSearchResultsModel = structureSearchModel.getResults();
		if (structureSearchModel != null && structureSearchResultsModel != null) {
			list = structureSearchResultsModel.getResults();
			System.out.println("结构搜索结果数量:" + list.size());
		}

		return list;
	}
	
	public static String formatString(String s) {

		if (s.contains(":") || s.contains("*") || s.contains("?") || s.contains("\"") || s.contains("|")
				|| s.contains("<") || s.contains(">") || s.contains(" ") || s.contains("\\") || s.contains("\\\\")
				|| s.contains("/")) {
			s = s.replace(" ", "");
			s = s.replace("\\\\", "");
			s = s.replace("\\", "");
			s = s.replace("/", "");
			s = s.replace("<", "");
			s = s.replace(">", "");

			s = s.replace("?", "");
			s = s.replace("\"", "");

			s = s.replace(":", "");
			s = s.replace("*", "");
			s = s.replace("|", "");
			s = s.replace(">", "");
		}
		return s;
	}

	/**
	 * 获取关联的板件
	 * 
	 * @param tcsession
	 * @param weldpoints
	 * @return
	 * @throws NotLoadedException
	 */
	public static HashMap<TCComponentBOMLine, TCComponent[]> getConnectedLines(TCSession tcsession,
			TCComponentBOMLine[] weldpoints) {
		StructureSearchService structuresearchservice = StructureSearchService.getService(tcsession);

		com.teamcenter.services.rac.manufacturing._2014_06.StructureSearch.SrchConnectedLinesResponse srchconnectedlinesresponse = structuresearchservice
				.searchConnectedLines(weldpoints);

		com.teamcenter.services.rac.manufacturing._2014_06.StructureSearch.SrchConnectedLinesOutput asrchconnectedlinesoutput[];
		int i1 = (asrchconnectedlinesoutput = srchconnectedlinesresponse.output).length;
		System.out.println("output.length:" + i1);
		HashMap<TCComponentBOMLine, TCComponent[]> map = new HashMap<TCComponentBOMLine, TCComponent[]>();

		String[] props = new String[] { "bl_DFL9SolItmPartRevision_dfl9_part_no" };

		for (int k = 0; k < i1; k++) {
			com.teamcenter.services.rac.manufacturing._2014_06.StructureSearch.SrchConnectedLinesOutput srchconnectedlinesoutput = asrchconnectedlinesoutput[k];
			TCComponent atccomponent1[] = srchconnectedlinesoutput.connectedLines;
			TCComponent connectioncomp = srchconnectedlinesoutput.connectionLine;
			TCComponentBOMLine connectionline = (TCComponentBOMLine) connectioncomp;

			LoadProperties(tcsession, atccomponent1, props);
			map.put(connectionline, atccomponent1);

		}
		return map;
	}

	public static void LoadProperties(TCSession session, TCComponent[] comps, String[] properties) {
		if (comps != null && comps.length > 0) {
			DataManagementService.getService(session).getProperties(comps, properties);
		}
	}

	public static void SetProperties(TCSession session, TCComponent[] comps, Map<String[], String[]> properties) {
		if (comps != null && comps.length > 0) {
			DataManagementService.getService(session).setProperties(comps, properties);
		}
	}

	// 正则表达式 : 完美
	public static boolean isNumber(String str) {
		if (str == null || str.isEmpty()) {
			return false;
		} else {
			//String reg = "^[0-9]+(.[0-9]+)?$";
//			String reg = "^-?[0-9]+.?[0-9]*";
//			return str.matches(reg);
			String pattern = "^[\\+\\-]?[\\d]+(\\.[\\d]+)?$";

			Pattern r = Pattern.compile(pattern);

			Matcher m = r.matcher(str);
			
			return m.matches();
		}
		// return false;
	}
	public static boolean regex(String content) {

		String pattern = "^[\\+\\-]?[\\d]+(\\.[\\d]+)?$";

		Pattern r = Pattern.compile(pattern);

		Matcher m = r.matcher(content);
		
		return m.matches();

	}

	/**
	 * 在制造工艺规划器中根据对象ID获取已打开的顶层BOMLine
	 * 
	 * @param rootItemId
	 * @return
	 */
	public static TCComponentBOMLine getOpenBOMLine(String rootItemId) {
		// System.out.println("-----------------getOpenBOMLine:"+rootItemId);
		try {
			AbstractAIFUIApplication localAbstractAIFUIApplication = AIFUtility.getActiveDesktop()
					.getCurrentApplication();
			if ((localAbstractAIFUIApplication instanceof AbstractBOMLineViewerApplication)) {
				// System.out.println("Hello
				// AbstractBOMLineViewerApplication!");
				AbstractBOMLineViewerApplication application = (AbstractBOMLineViewerApplication) localAbstractAIFUIApplication;
				AbstractViewableTreeTable[] treeTables = application.getViewableTreeTables();
				if (treeTables == null) {
					System.out.println("treeTables==null");
				} else {
					// System.out.println("treeTables length=" +
					// treeTables.length);
					for (int i = 0; i < treeTables.length; i++) {
						AbstractViewableTreeTable treeTable = treeTables[i];
						BOMLineNode node = treeTable.getRootBOMLineNode();
						if (node == null) {
							continue;
						}
						TCComponentBOMLine rootBOMLine = node.getBOMLine();
						// System.out.println("rootBOMLine="+
						// rootBOMLine.toString());

						String id = rootBOMLine.getItem().getProperty("item_id");
						if (rootItemId.equals(id)) {
							// System.out.println("找到BOMLine");
							return rootBOMLine;
						}
					}
				}
			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null;
	}

	/**
	 * 调用开启旁路功能
	 * 
	 * @param session
	 * @param flag
	 * @return
	 * @throws TCException
	 */
	public static boolean callByPass(TCSession session, boolean flag) throws TCException {
//		TCUserService service = session.getUserService();
//		service.call("tc_set_bypass", new Object[] { flag });
		SetByPassService byPassService = SetByPassService.getService(session.getSoaConnection());
		byPassService.bypass(flag);
		return true;
	}

	public static String[][] getAllProperties(TCSession session, TCComponent[] comps, String[] properties)
			throws TCException {
		ServiceData serviceData = DataManagementService.getService(session).getProperties(comps, properties);
		if (serviceData.sizeOfPartialErrors() <= 0) {
			int sizea = serviceData.sizeOfPlainObjects();
			int count = serviceData.sizeOfPlainObjects();;
			//int count = comps.length;
			int vcount = properties.length;
			TCComponent comp;
			String[][] values = new String[sizea][vcount];
			try {
				for (int i = 0; i < count; i++) {
					values[i] = new String[vcount];
					comp = serviceData.getPlainObject(i);

					for (int j = 0; j < vcount; j++) {
						try {
							values[i][j] = comp.getPropertyDisplayableValue(properties[j]);
						} catch (NotLoadedException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						} // getProperty(properties[j]);
					}
				}
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

			return values;
		} else {
			System.out.println("ERROR:LoadProperties error!");
		}
		return null;
	}

	public static void setCompsProperty(TCSession session, TCComponent[] comps, String prop, String value) {
		System.out.println(">>>setCompsProperty");
		if (comps != null && comps.length > 0) {
			// LoadProperties( session,comps,new String[] {prop});
			DataManagementService dataManagementService = DataManagementService.getService(session);
			HashMap hashMap = new HashMap();
			DataManagement.VecStruct vecStruct1 = new DataManagement.VecStruct();
			vecStruct1.stringVec = new String[1];
			vecStruct1.stringVec[0] = value;
			hashMap.put(prop, vecStruct1);
			dataManagementService.setProperties(comps, hashMap);
			dataManagementService.refreshObjects(comps);

		}
	}

	public static void setAllCompsProperty(TCSession session, Map<TCComponentBOMLine, String[]> map, String[] props)
			throws TCException {
		System.out.println(">>>setCompsProperty");
		if (map != null && map.size() > 0) {
			for (Map.Entry<TCComponentBOMLine, String[]> entry : map.entrySet()) {
				TCComponentBOMLine bl = entry.getKey();
				String[] value = entry.getValue();
				TCComponentItemRevision rev = bl.getItemRevision();
//				TCComponent[] form = rev.getRelatedComponents("TC_Feature_Form_Relation");
//				if (form != null && form.length > 0) {
//					form[0].lock();
//					form[0].setProperties(props, value);
//					form[0].save();
//					form[0].unlock();
//				}
				bl.lock();
				bl.setProperties(props, value);
				bl.save();
				bl.unlock();
			}
		}
	}

	/*
	 * ****************************** 根据首选项查询数据集
	 */
	public static InputStream getReportTempByprefercen(TCSession session, String prefrencename, int index) {

		InputStream inputStream = null;
		try {
			File file = null;
			TCPreferenceService preferenceService = session.getPreferenceService();

			String str = preferenceService.getPreferenceDescription(prefrencename);
			if (str != null) {
				String[] values = preferenceService.getStringValues(prefrencename);
				if (values != null && values.length >= index) {
					String value = values[index - 1];
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
							inputStream = new FileInputStream(file);
						}
					}
				}
			}
			return inputStream;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return inputStream;
	}

	// 把生成的报表，作为数据集放到指定对象下
	public static void saveFilesToFolder(TCSession session, TCComponent tcc, String datasetname, String filename,
			String filetype, String type) {
		try {
			String fullFileName = FileUtil.getReportFileName(filename);
			TCComponentUser user = session.getUser();
			TCComponentFolder folder = (TCComponentFolder) tcc;

			TCComponentItemType tcccomponentitemtype = (TCComponentItemType) session.getTypeComponent(filetype);

			TCComponentItem tccomponentitem = tcccomponentitemtype.create("", "", filetype, datasetname, "desc", null);
			tccomponentitem.setProperty("b8_BIWProcDocType", type);
			tccomponentitem.lock();
			tccomponentitem.save();
			tccomponentitem.unlock();
			folder.add("contents", tccomponentitem);
			TCComponentItemRevision rev = tccomponentitem.getLatestItemRevision();
			TCComponentDataset ds = Util.createDataset(session, datasetname, fullFileName, "MSExcelX", "excel");
			// 添加文档与数据集的关系
			rev.add("IMAN_specification", ds);
			rev.lock();
			rev.save();
			rev.unlock();
			// 删除中间文件
			File file = new File(fullFileName);
			if (file.isFile()) {
				file.delete();
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/*
	 * 根据首选项获取车型代号
	 */
	public static String getDFLProjectIdVehicle(String FamlilyCode) {
		// 读取 项目-车型 首选项
		String VehicleNo = "";
		Map<String, String> projVehMap = ReportUtils.getDFL_Project_VehicleNo();

		if (projVehMap.size() < 1) {
			VehicleNo = FamlilyCode;
		} else {
			VehicleNo = projVehMap.get(FamlilyCode);
			if (VehicleNo == null) {
				if (FamlilyCode != null) {
					VehicleNo = FamlilyCode;
				}
			}
		}
		return VehicleNo;
	}

	public static boolean isContinue(String message) {
		int i = ConfirmationDialog.post("信息", message, false);
		// TODO Auto-generated method stub
		System.out.println("i=" + i);
		switch (i) {
		case 2:
			return true;
		case 1:
			return false;
		case 3:
			return false;
		}
		return false;
	}

	public static boolean getIsVirtualLine(TCComponentBOMLine topbomline) {
		// TODO Auto-generated method stub
		// 根据产线下是否存在产线判断是否为虚层
		boolean flag = false;
		ArrayList list = Util.getChildrenByParent(topbomline);
		if (list != null) {
			for (int i = 0; i < list.size(); i++) {
				TCComponentBOMLine bl = (TCComponentBOMLine) list.get(i);
				try {
					TCComponentItemRevision rev = bl.getItemRevision();
					if (rev.isTypeOf("B8_BIWMEProcLineRevision")) {
						flag = true;
					}
				} catch (TCException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
		return flag;
	}

	/**
	 * 获取所选虚层集合
	 * 
	 * @param session
	 * @param comp
	 * @return
	 * @throws Exception
	 */
	public static List getChildrenByParent(InterfaceAIFComponent[] aifc) throws Exception {
		List list = new ArrayList();
		for (InterfaceAIFComponent aif : aifc) {
			TCComponentBOMLine parentbl = (TCComponentBOMLine) aif;
			TCComponentItemRevision parentrev = parentbl.getItemRevision();
			// 如果是顶层，遍历获取
			if (parentrev.isTypeOf("B8_BIWPlantBOPRevision")) {
				AIFComponentContext[] childrens = parentbl.getChildren();
				for (AIFComponentContext chil : childrens) {
					TCComponentBOMLine bl = (TCComponentBOMLine) chil.getComponent();
					boolean flag = Util.getIsVirtualLine(bl); // 判断是否为虚层
					if (flag) {
						if (!list.contains(bl)) {
							list.add(bl);
						}
					}
				}
			} else {// 如果是虚层，直接添加
				if (!list.contains(parentbl)) {
					list.add(parentbl);
				}
			}

		}
		return list;
	}

	public static File getRCPPluginInsideFile(String name) {

		String tempPath = System.getProperty("java.io.tmpdir");
		InputStream inputStream = Util.class.getResourceAsStream(name);
		if (tempPath.endsWith("\\")) {
			tempPath = tempPath.substring(0, tempPath.length() - 1);
		}
		//String filePath = tempPath + "\\" + name;
		
		String filePath = "D:" + "\\" + name;
		File file = new File(filePath);

		if (file.exists()) {
			file.delete();
		}

		try {
			FileOutputStream fileOutputStream = new FileOutputStream(file);

			byte[] b = new byte[1024 * 5];
			int len;
			while ((len = inputStream.read(b)) != -1) {
				fileOutputStream.write(b, 0, len);
			}
			fileOutputStream.flush();
			fileOutputStream.close();
			inputStream.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return file;
	}

	// 判断焊装产线工艺下是否有多个工位
	public static boolean getIsMEProcStat(TCComponentBOMLine bl) {
		// TODO Auto-generated method stub
		boolean flag = false;
		try {
			AIFComponentContext[] children = bl.getChildren();
            int count =0;
			for (AIFComponentContext chil : children) {
				TCComponentItemRevision rev = ((TCComponentBOMLine) chil.getComponent()).getItemRevision();
				if (rev.isTypeOf("B8_BIWMEProcStatRevision")) {
					count++;
				}
			}
			if(count>1) {
				flag = true;
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return flag;
	}

	
	//将文档指派项目
	public static void assignProjectComp(TCComponentItemRevision oldrev, TCComponent[] projects) throws TCException {
		// TODO Auto-generated method stub
		if(projects!=null && projects.length>0) {
			TCComponentItemRevision[] tccitemrev = {oldrev};
			for(TCComponent tcc : projects) {
				if(tcc instanceof TCComponentProject) {
					TCComponentProject pro = (TCComponentProject) tcc;
					pro.assignToProject(tccitemrev);
				}
			}
		}
	}

	/*
	 * 根据对象类型获取对象显示名称
	 */
	public static String getObjectDisplayName(TCSession session,String objecttype) {
		String dispalyname = "";
		try {
			TCComponentItemType tcccomponentitemtype = (TCComponentItemType) session.getTypeComponent(objecttype);
			dispalyname = tcccomponentitemtype.getDisplayType();
		} catch (TCException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		return dispalyname;				
	}		

	/**
	 * 比较带X通配符的字符串是否一致
	 * @param sorceStr
	 * @param targetStr
	 * @return
	 */
    public static boolean getIsEqueal(String sorceStr,String targetStr)
    {
    	boolean flag = true;
    	try {
			if(sorceStr==null || sorceStr.isEmpty() || targetStr==null || targetStr.isEmpty())
			{
				return false;
			}
			if(sorceStr.length()<targetStr.length())
			{
				return false;
			}
			for(int i=0;i<targetStr.length();i++)
			{
				if(targetStr.charAt(i)!='X' && targetStr.charAt(i)!=sorceStr.charAt(i))
				{
					flag = false;
					break;
				}
			}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    	return flag;
    }
}
