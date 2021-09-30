package com.dfl.report.common;

import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AifrcpPlugin;
import com.teamcenter.rac.kernel.TCAccessControlService;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentEnvelope;
import com.teamcenter.rac.kernel.TCComponentEnvelopeType;
import com.teamcenter.rac.kernel.TCComponentForm;
import com.teamcenter.rac.kernel.TCComponentFormType;
import com.teamcenter.rac.kernel.TCComponentGroup;
import com.teamcenter.rac.kernel.TCComponentGroupMember;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentItemType;
import com.teamcenter.rac.kernel.TCComponentListOfValues;
import com.teamcenter.rac.kernel.TCComponentTcFile;
import com.teamcenter.rac.kernel.TCComponentUser;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCProperty;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.schemas.soa._2006_03.exceptions.ServiceException;
import com.teamcenter.services.internal.rac.core.ICTService;
import com.teamcenter.services.rac.core.DataManagementService;
import com.teamcenter.services.rac.core._2006_03.DataManagement.ObjectOwner;
import com.teamcenter.services.rac.core._2006_03.DataManagement.Relationship;
import com.teamcenter.services.rac.core._2008_06.DataManagement;
import com.teamcenter.services.rac.core._2015_07.DataManagement.CreateIn2;
import com.teamcenter.services.rac.core._2015_07.DataManagement.CreateInput2;
import com.teamcenter.soa.client.model.ErrorStack;
import java.awt.image.BufferedImage;
import java.io.PrintStream;
import java.io.UnsupportedEncodingException;
import java.math.BigInteger;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class TCComponentUtils {
	private static final String VIEWER_VIEW = "com.teamcenter.rac.ui.views.ViewerView";
	private static final String SUMMARY_VIEW = "com.teamcenter.rac.ui.views.SummaryView";
	private static final String PERSPECTIVE_MYTEAMCENTER = "com.teamcenter.rac.ui.perspectives.navigatorPerspective";
	public static final String FORM_ITEM_MASTER = "Item Master";
	public static final String IMAN_MASTER_FORM = "IMAN_master_form";
	public static final String FORM_ITEM_REVISION_MASTER = "ItemRevision Master";
	public static final String IMAN_MASTER_FORM_REV = "IMAN_master_form_rev";
	public static final String RELATION_PART_DESIGN = "TC_Is_Represented_By";
	public static final String PROPERTY_DESIGN_PART = "representation_for";
	public static final String REL_PROBLEM = "CMHasProblemItem";
	public static final String REL_IMPACTED = "CMHasImpactedItem";
	public static final String REL_SOLUTION = "CMHasSolutionItem";
	public static final String REL_REFERENCE = "CMReferences";
	public static final String REL_IMPLEMENT = "CMImplements";
	public static final String REL_PLAN = "CMHasWorkBreakdown";
	private static TCSession session = TCUtils.getTCSession();

	public static void openWithMyTc(String cmpUid) {
		TCComponent cmp = loadObject(cmpUid);
		if (cmp != null)
			openWithMyTc(cmp);
	}

	public static void openWithMyTc(TCComponent component) {
		AifrcpPlugin.getDefault().openPerspective("com.teamcenter.rac.ui.perspectives.navigatorPerspective");
		AifrcpPlugin.getDefault().openComponents("com.teamcenter.rac.ui.perspectives.navigatorPerspective",
				new InterfaceAIFComponent[] { component });
	}

	public static String getOwnerUserId(TCComponent component) throws TCException {
		if (component != null) {
			TCProperty property = component.getTCProperty("owning_user");
			if (property != null) {
				TCComponent owningUser = property.getReferenceValue();
				if (owningUser != null) {
					return ((TCComponentUser) owningUser).getUserId();
				}
			}
		}
		return null;
	}

	public static TCComponentItemType getItemType(String type) throws TCException {
		return (TCComponentItemType) TCUtils.getTypeComponent(type);
	}

//  public static boolean isItemExist(String itemID)
//    throws TCException
//  {
//    TCComponent[] result = TCComponentQueryUtils.queryItemById(itemID);
//    return (result != null) && (result.length > 0);
//  }

	public static boolean isRevisionExist(TCComponentItem item, String strVer) throws TCException {
		TCComponentItemRevision[] workingRevs = item.getWorkingItemRevisions();
		for (int i = 0; i < workingRevs.length; i++) {
			String strTempVer = workingRevs[i].getTCProperty("item_revision_id").getStringValue();
			if (strTempVer.equals(strVer)) {
				return true;
			}
		}

		TCComponentItemRevision[] inprocessRevs = item.getInProcessItemRevisions();
		for (int i = 0; i < inprocessRevs.length; i++) {
			String strTempVer = inprocessRevs[i].getTCProperty("item_revision_id").getStringValue();
			if (strTempVer.equals(strVer)) {
				return true;
			}
		}

		TCComponentItemRevision[] releasedRevs = item.getReleasedItemRevisions();
		for (int i = 0; i < releasedRevs.length; i++) {
			String strTempVer = releasedRevs[i].getTCProperty("item_revision_id").getStringValue();
			if (strTempVer.equals(strVer)) {
				return true;
			}
		}

		return false;
	}

//  public static TCComponent getItemById(String itemId)
//    throws TCException
//  {
//    TCComponent[] result = TCComponentQueryUtils.queryByProperty(TCComponentQueryUtils.QUERY_ITEM, "item_id", 
//      itemId);
//    if ((result != null) && (result.length > 0)) {
//      return result[0];
//    }
//    return null;
//  }

	public static TCComponent findItem(String type, String id) throws TCException {
		TCComponentItemType itemType = getItemType(type);
		TCComponent[] items = itemType.findItems(id);
		if ((items != null) && (items.length > 0)) {
			return items[0];
		}
		return null;
	}

//  public static TCComponent getComponentByIdAndType(String partId, String partType)
//    throws TCException
//  {
//    if ((partId == null) || (partType == null)) {
//      return null;
//    }
//
//    TCComponent[] components = TCComponentQueryUtils.queryItemById(partId);
//    if ((components != null) && (components.length > 0)) {
//      for (TCComponent component : components) {
//        if (component.getType().equals(partType)) {
//          return component;
//        }
//      }
//    }
//
//    return null;
//  }

	public static String generateNewID(String itemType) throws TCException {
		TCComponentItemType type = getItemType(itemType);
		return type.getNewID();
	}

	public static String generateNewRevID(TCComponentItem item) throws TCException {
		TCComponentItemType type = getItemType(item.getType());
		return type.getNewRev(item);
	}

	public static TCComponent[] findItems(String type, String id) throws TCException {
		TCComponentItemType itemType = getItemType(type);
		return itemType.findItems(id);
	}

	public static TCComponentItem create(String type, String id, String rev, String name, String description,
			TCComponent uom) throws TCException {
		TCComponentItemType componentItemType = getItemType(type);
		TCComponentItem item = componentItemType.create(id, rev, type, name, description, uom);
		return item;
	}

	public static TCComponentItem create(String type, String rev, String name, String description, TCComponent uom)
			throws TCException {
		TCComponentItemType componentItemType = getItemType(type);
		TCComponentItem item = componentItemType.create("", rev, type, name, description, uom);
		return item;
	}

	public static DataManagement.CreateResponse create(Map<String, Object> itemMap, Map<String, Object> itemRevisionMap,
			Map<String, Object> itemRevMasterFormMap) throws ServiceException {
		DataManagementService dmService = DataManagementService.getService(session);

		DataManagement.CreateInput itemRevMasterFormDef = createInputForm(itemRevMasterFormMap);

		DataManagement.CreateInput itemRevisionDef = createInputItemRevision(itemRevisionMap, itemRevMasterFormDef);

		DataManagement.CreateIn itemDef = createInputItem(itemMap, itemRevisionDef);

		DataManagement.CreateResponse createObjResponse = dmService
				.createObjects(new DataManagement.CreateIn[] { itemDef });

		return createObjResponse;
	}

	public static void updateReferenceProperty(TCComponent component, String key, TCComponent value)
			throws TCException {
		if ((key != null) && (!"".equals(key)) && (value != null)) {
			TCProperty property = component.getTCProperty(key);

			if ((property != null) && (property.isReferenceType()) && (property.isNotArray()))
				component.setReferenceProperty(key, value);
		}
	}

	public static void updateProperties(TCComponent component, TCProperty[] tcPropertyArray) throws TCException {
		if ((component != null) && (tcPropertyArray != null))
			component.setTCProperties(tcPropertyArray);
	}

	/** @deprecated */
	public static String getObjectDisplay(TCComponent component) throws TCException {
		if ((component instanceof TCComponentItem)) {
			return component.getStringProperty("item_id");
		}

		if ((component instanceof TCComponentItemRevision)) {
			return component.getStringProperty("item_id") + "/" + component.getStringProperty("item_revision_id");
		}

		return null;
	}

	public static boolean isTcProperty(TCComponent component, String propName) throws TCException {
		if ((component != null) && (propName != null) && (!"".equals(propName))) {
			TCProperty property = component.getTCProperty(propName);
			if (property != null) {
				return true;
			}
		}
		return false;
	}

	public static void createRelation(TCComponent primary, TCComponent secondary, String relationName)
			throws TCException {
		TCComponent[] secondarys = primary.getRelatedComponents(relationName);
		boolean isExist = false;
		for (TCComponent comp : secondarys) {
			if (comp == secondary) {
				isExist = true;
			}
		}

		if (!isExist)
			primary.add(relationName, secondary);
	}

	public static void removeSecondarysByRelation(TCComponent primary, String relationName) throws TCException {
		if (primary != null) {
			TCComponent[] secondarys = primary.getRelatedComponents(relationName);
			if ((secondarys != null) && (secondarys.length > 0))
				primary.cutOperation(relationName, secondarys);
		}
	}

	public static boolean checkPrivilege(TCComponent tcComponent, String privilege) throws TCException {
		tcComponent.refresh();
		return session.getTCAccessControlService().checkPrivilege(tcComponent, privilege);
	}

	public static TCComponentForm getMasterForm(TCComponent target) throws TCException {
		if (target != null) {
			if ((target instanceof TCComponentItem)) {
				return (TCComponentForm) ((TCComponentItem) target).getRelatedComponent("IMAN_master_form");
			}

			if ((target instanceof TCComponentItemRevision)) {
				return (TCComponentForm) ((TCComponentItemRevision) target).getRelatedComponent("IMAN_master_form_rev");
			}
		}

		return null;
	}

	public static TCComponentFormType getFormType(String formType) throws TCException {
		return (TCComponentFormType) TCUtils.getTypeComponent(formType);
	}

	public static TCComponentForm create(String name, String desc, String type, boolean saveDB) throws TCException {
		TCComponentFormType componentFormType = getFormType(type);
		return componentFormType.create(name, desc, type, saveDB);
	}

//  public static TCComponentEnvelope createEnvelope(String subject, String comments)
//    throws TCException, UnsupportedEncodingException
//  {
//    if ((subject != null) && (!"".equals(subject))) {
//      TCComponentEnvelopeType tccomponentenvelopetype = (TCComponentEnvelopeType)session
//        .getTypeComponent("Envelope");
//      if (subject.length() > 120) {
//        subject = BaseUtils.subStringByByte(subject, 120) + "...";
//      }
//      if ((comments != null) && (comments.length() > 140)) {
//        comments = BaseUtils.subStringByByte(comments, 140) + "...";
//      }
//
//      return tccomponentenvelopetype.create(subject, comments, "Envelope");
//    }
//    return null;
//  }

	public static TCComponentItemRevision revise(TCComponentItemRevision selectedRevision) throws TCException {
		return selectedRevision.saveAs("");
	}

	public static TCComponentItemRevision getLatestItemRevision(TCComponentItemRevision target) throws TCException {
		if (target != null) {
			TCComponentItem item = target.getItem();
			return item.getLatestItemRevision();
		}
		return null;
	}

	public static TCComponentItemRevision[] getAllRevision(TCComponentItem item) throws Exception {
		TCComponentItemRevision[] workingRevs = item.getWorkingItemRevisions();
		TCComponentItemRevision[] inprocessRevs = item.getInProcessItemRevisions();
		TCComponentItemRevision[] releasedRevs = item.getReleasedItemRevisions();

		int count = workingRevs.length + inprocessRevs.length + releasedRevs.length;

		TCComponentItemRevision[] revs = new TCComponentItemRevision[count];

		int top = 0;
		for (int i = 0; i < workingRevs.length; i++) {
			revs[top] = workingRevs[i];
			top++;
		}

		for (int i = 0; i < inprocessRevs.length; i++) {
			revs[top] = inprocessRevs[i];
			top++;
		}

		for (int i = 0; i < releasedRevs.length; i++) {
			revs[top] = releasedRevs[i];
			top++;
		}

		return revs;
	}

	public TCComponentItemRevision getRevision(TCComponentItem item, String revNo) throws TCException {
		TCComponentItemRevision[] workingRevs = item.getWorkingItemRevisions();
		for (int i = 0; i < workingRevs.length; i++) {
			String strTempVer = workingRevs[i].getTCProperty("item_revision_id").getStringValue();
			if (strTempVer.equals(revNo)) {
				return workingRevs[i];
			}
		}

		TCComponentItemRevision[] inprocessRevs = item.getInProcessItemRevisions();
		for (int i = 0; i < inprocessRevs.length; i++) {
			String strTempVer = inprocessRevs[i].getTCProperty("item_revision_id").getStringValue();
			if (strTempVer.equals(revNo)) {
				return inprocessRevs[i];
			}
		}

		TCComponentItemRevision[] releasedRevs = item.getReleasedItemRevisions();
		for (int i = 0; i < releasedRevs.length; i++) {
			String strTempVer = releasedRevs[i].getTCProperty("item_revision_id").getStringValue();
			if (strTempVer.equals(revNo)) {
				return releasedRevs[i];
			}
		}

		return null;
	}

	public static TCComponentItemRevision getMaxItemRevision(TCComponentItemRevision target) throws TCException {
		List revList = new ArrayList();
		TCComponentItem item = target.getItem();

		TCComponentItemRevision[] revs1 = item.getWorkingItemRevisions();
		TCComponentItemRevision[] revs2 = item.getInProcessItemRevisions();
		TCComponentItemRevision[] revs3 = item.getReleasedItemRevisions();

		for (TCComponentItemRevision rev : revs1) {
			revList.add(rev);
		}
		for (TCComponentItemRevision rev : revs2) {
			revList.add(rev);
		}
		for (TCComponentItemRevision rev : revs3) {
			revList.add(rev);
		}

		Collections.sort(revList, new Comparator() {
			public int compare(Object a, Object b) {
				try {
					String itemRevisionId1 = ((TCComponentItemRevision) a).getStringProperty("item_revision_id");
					String itemRevisionId2 = ((TCComponentItemRevision) b).getStringProperty("item_revision_id");
					return itemRevisionId2.compareTo(itemRevisionId1);
				} catch (TCException e) {
					e.printStackTrace();
				}
				return 0;
			}
		});
		if ((revList != null) && (revList.size() > 0)) {
			return (TCComponentItemRevision) revList.get(0);
		}
		return null;
	}

	public static TCComponentItem saveAs(TCComponentItem item) throws TCException {
		return saveAs(null, item);
	}

	public static TCComponentItem saveAs(String itemId, TCComponentItem item) throws TCException {
		return saveAs(itemId, item.getLatestItemRevision());
	}

	public static TCComponentItem saveAs(TCComponentItemRevision revision) throws TCException {
		return saveAs(null, revision);
	}

	public static TCComponentItem saveAs(String itemId, TCComponentItemRevision revision) throws TCException {
		return revision.saveAsItem(itemId, "A");
	}

	public static void deleteItem(TCComponent item) throws TCException {
		item.delete();
	}

	public static TCComponent getComponentByUid(String uid) throws TCException {
		return session.stringToComponent(uid);
	}

	public static TCComponent[] getCompsByRelation(TCComponent primary, String relation) throws TCException {
		return primary.getRelatedComponents(relation);
	}

	public static TCComponent[] getSecondarysByRelationAndType(TCComponent primary, String relation, String[] types)
			throws TCException {
		Arrays.sort(types);
		List result = new ArrayList();
		TCComponent[] secondarys = primary.getRelatedComponents(relation);
		if (secondarys != null) {
			for (TCComponent secondary : secondarys) {
				if (Arrays.binarySearch(types, secondary.getType()) > -1) {
					result.add(secondary);
				}
			}
		}
		if (result.size() > 0) {
			TCComponent[] finalResult = new TCComponent[result.size()];
			result.toArray(finalResult);
			return finalResult;
		}
		return null;
	}

	public static TCComponent[] getSecondaryByType(TCComponent primary, String[] types) throws TCException {
		Arrays.sort(types);
		List result = new ArrayList();
		TCComponent[] secondarys = primary.getRelatedComponents();
		if (secondarys != null) {
			for (TCComponent secondary : secondarys) {
				if (Arrays.binarySearch(types, secondary.getType()) > -1) {
					result.add(secondary);
				}
			}
		}
		if (result.size() > 0) {
			TCComponent[] finalResult = new TCComponent[result.size()];
			result.toArray(finalResult);
			return finalResult;
		}
		return null;
	}

	public static TCComponent[] getDesignsByPart(TCComponent part) throws TCException {
		TCComponentItemRevision partRevision = null;

		if ((part instanceof TCComponentItem)) {
			partRevision = ((TCComponentItem) part).getLatestItemRevision();
		}

		if ((part instanceof TCComponentItemRevision)) {
			partRevision = (TCComponentItemRevision) part;
		}

		if (partRevision != null) {
			TCProperty relationProperty = partRevision.getTCProperty("TC_Is_Represented_By");
			if (relationProperty != null) {
				return getCompsByRelation(partRevision, "TC_Is_Represented_By");
			}
		}

		return null;
	}

	public static void relatedDesignToPart(TCComponentItemRevision partRev, TCComponentItemRevision designRev)
			throws TCException {
		if ((partRev != null) && (designRev != null))
			createRelation(partRev, designRev, "TC_Is_Represented_By");
	}

	public static TCComponent[] getPartsByDesign(TCComponent design) throws TCException {
		TCComponentItemRevision designRevision = null;

		if ((design instanceof TCComponentItem)) {
			designRevision = ((TCComponentItem) design).getLatestItemRevision();
		}

		if ((design instanceof TCComponentItemRevision)) {
			designRevision = (TCComponentItemRevision) design;
		}

		if (designRevision != null) {
			designRevision.refresh();
			TCProperty relationProperty = designRevision.getTCProperty("representation_for");
			if (relationProperty != null) {
				return getCompsByRelation(designRevision, "representation_for");
			}
		}

		return null;
	}

	/** @deprecated */
	public static String getComponentBrief(TCComponent component) throws TCException {
		if (component != null) {
			return component.getStringProperty("item_id") + "/" + component.getStringProperty("item_revision_id") + "-"
					+ component.getStringProperty("object_name");
		}
		return null;
	}

	public static TCComponent[] getPrimary(TCComponent component) throws TCException {
		if (component != null) {
			AIFComponentContext[] aifComps = component.getPrimary();
			TCComponent[] primaryComps = new TCComponent[aifComps.length];
			for (int i = 0; i < aifComps.length; i++) {
				primaryComps[i] = ((TCComponent) aifComps[i].getComponent());
			}
			if (primaryComps.length > 0) {
				return primaryComps;
			}
		}
		return null;
	}

	public static List<String> serviceDataError(DataManagement.CreateResponse response) {
		if (response != null) {
			com.teamcenter.soa.client.model.ServiceData data = response.serviceData;
			if ((data != null) && (data.sizeOfPartialErrors() > 0)) {
				List msgs = new ArrayList();
				for (int i = 0; i < data.sizeOfPartialErrors(); i++) {
					for (String msg : data.getPartialError(i).getMessages()) {
						msgs.add(msg);
					}
				}
				return msgs;
			}
		}
		return null;
	}

	private static DataManagement.CreateIn createInputItem(Map<String, Object> itemMap,
			DataManagement.CreateInput itemRevisionDef) {
		DataManagement.CreateIn itemDef = new DataManagement.CreateIn();
		for (String key : itemMap.keySet()) {
			Object value = itemMap.get(key);
			if ("item_id".equals(key)) {
				itemDef.data.stringProps.put(key, value);
			} else if ("object_type".equals(key)) {
				itemDef.data.boName = value.toString();
			} else if ((value instanceof String)) {
				itemDef.data.stringProps.put(key, value);
			} else if ((value instanceof Integer)) {
				itemDef.data.intProps.put(key, BigInteger.valueOf(((Integer) value).intValue()));
			} else if ((value instanceof BigInteger)) {
				itemDef.data.intProps.put(key, (BigInteger) value);
			} else if ((value instanceof Double)) {
				itemDef.data.doubleProps.put(key, (Double) value);
			} else if ((value instanceof Boolean)) {
				itemDef.data.boolProps.put(key, (Boolean) value);
			} else if ((value instanceof Date)) {
				Calendar calendar = Calendar.getInstance();
				calendar.setTime((Date) value);
				itemDef.data.dateProps.put(key, calendar);
			} else if ((value instanceof Calendar)) {
				itemDef.data.dateProps.put(key, (Calendar) value);
			} else if ((value instanceof GregorianCalendar)) {
				itemDef.data.dateProps.put(key, (GregorianCalendar) value);
			} else if ((value instanceof TCComponent)) {
				itemDef.data.tagProps.put(key, (TCComponent) value);
			}
		}

		itemDef.data.compoundCreateInput.put("revision", new DataManagement.CreateInput[] { itemRevisionDef });
		return itemDef;
	}

	private static DataManagement.CreateInput createInputItemRevision(Map<String, Object> itemRevisionMap,
			DataManagement.CreateInput itemRevMasterFormDef) {
		DataManagement.CreateInput itemRevisionDef = new DataManagement.CreateInput();

		for (String key : itemRevisionMap.keySet()) {
			Object value = itemRevisionMap.get(key);
			if ("object_type".equals(key)) {
				itemRevisionDef.boName = value.toString();
			} else if ((value instanceof String[])) {
				itemRevisionDef.stringArrayProps.put(key, (String[]) value);
			} else if ((value instanceof String)) {
				itemRevisionDef.stringProps.put(key, value);
			} else if ((value instanceof Integer)) {
				itemRevisionDef.intProps.put(key, BigInteger.valueOf(((Integer) value).intValue()));
			} else if ((value instanceof BigInteger)) {
				itemRevisionDef.intProps.put(key, (BigInteger) value);
			} else if ((value instanceof Double)) {
				itemRevisionDef.doubleProps.put(key, (Double) value);
			} else if ((value instanceof Boolean)) {
				itemRevisionDef.boolProps.put(key, (Boolean) value);
			} else if ((value instanceof Date)) {
				Calendar calendar = Calendar.getInstance();
				calendar.setTime((Date) value);
				itemRevisionDef.dateProps.put(key, calendar);
			} else if ((value instanceof Calendar)) {
				itemRevisionDef.dateProps.put(key, (Calendar) value);
			} else if ((value instanceof GregorianCalendar)) {
				itemRevisionDef.dateProps.put(key, (GregorianCalendar) value);
			} else if ((value instanceof TCComponent)) {
				itemRevisionDef.tagProps.put(key, (TCComponent) value);
			}

		}

		if (itemRevMasterFormDef != null) {
			itemRevisionDef.compoundCreateInput.put("IMAN_master_form_rev",
					new DataManagement.CreateInput[] { itemRevMasterFormDef });
		}

		return itemRevisionDef;
	}

	private static DataManagement.CreateInput createInputForm(Map<String, Object> valueMap) {
		if (valueMap != null) {
			DataManagement.CreateInput formDef = new DataManagement.CreateInput();
			for (String key : valueMap.keySet()) {
				Object value = valueMap.get(key);
				if ("object_type".equals(key)) {
					formDef.boName = value.toString();
				} else if ((value instanceof String)) {
					formDef.stringProps.put(key, value);
				} else if ((value instanceof Integer)) {
					formDef.intProps.put(key, BigInteger.valueOf(((Integer) value).intValue()));
				} else if ((value instanceof BigInteger)) {
					formDef.intProps.put(key, (BigInteger) value);
				} else if ((value instanceof Boolean)) {
					formDef.boolProps.put(key, (Boolean) value);
				} else if ((value instanceof Date)) {
					Calendar calendar = Calendar.getInstance();
					calendar.setTime((Date) value);
					formDef.dateProps.put(key, calendar);
				} else if ((value instanceof Calendar)) {
					formDef.dateProps.put(key, (Calendar) value);
				} else if ((value instanceof GregorianCalendar)) {
					formDef.dateProps.put(key, (GregorianCalendar) value);
				} else if ((value instanceof TCComponent)) {
					formDef.tagProps.put(key, (TCComponent) value);
				}
			}

			return formDef;
		}
		return null;
	}

	public static TCComponentItem getItem(DataManagement.CreateResponse response) throws ServiceException {
		if (response != null) {
			DataManagement.CreateOut[] outputs = response.output;
			if (outputs != null) {
				for (DataManagement.CreateOut out : outputs) {
					if (out.objects != null) {
						for (TCComponent obj : out.objects) {
							if ((obj instanceof TCComponentItem)) {
								return (TCComponentItem) obj;
							}
						}
					}
				}
			}
		}
		return null;
	}

	public static TCComponent create(String componentType, Map<String, String> stringProperties,
			Map<String, BigInteger> intProperties, Map<String, TCComponent> tagProperties) throws ServiceException {
		TCSession session = TCUtils.getTCSession();
		DataManagementService ds = DataManagementService.getService(session);

		DataManagement.CreateIn in = new DataManagement.CreateIn();
		in.clientId = (componentType + System.currentTimeMillis());

		DataManagement.CreateInput input = new DataManagement.CreateInput();
		input.boName = componentType;
		if (stringProperties != null) {
			input.stringProps = stringProperties;
		}
		if (intProperties != null) {
			input.intProps = intProperties;
		}
		if (tagProperties != null) {
			input.tagProps = tagProperties;
		}
		in.data = input;

		DataManagement.CreateResponse resp = ds.createObjects(new DataManagement.CreateIn[] { in });
		if (resp.serviceData.sizeOfPartialErrors() > 0) {
			throw new ServiceException(Arrays.toString(resp.serviceData.getPartialError(0).getMessages()));
		}
		return resp.serviceData.getCreatedObject(0);
	}

	public static TCComponent create(String componentType, Map<String, String> itemStringMap,
			Map<String, TCComponent> itemTagMap, Map<String, String> revStringMap, Map<String, Double> revDoubleMap,
			Map<String, TCComponent> revTagMap) throws ServiceException {
		TCSession session = TCUtils.getTCSession();
		DataManagementService ds = DataManagementService.getService(session);

		DataManagement.CreateInput revInput = new DataManagement.CreateInput();
		revInput.boName = (componentType + "Revision");
		revInput.stringProps = revStringMap;
		revInput.doubleProps = revDoubleMap;

		DataManagement.CreateIn in = new DataManagement.CreateIn();
		in.clientId = (componentType + System.currentTimeMillis());
		DataManagement.CreateInput input = new DataManagement.CreateInput();
		input.boName = componentType;
		input.stringProps = itemStringMap;
		input.tagProps = itemTagMap;
		in.data = input;
		in.data.compoundCreateInput.put("revision", new DataManagement.CreateInput[] { revInput });

		DataManagement.CreateResponse resp = ds.createObjects(new DataManagement.CreateIn[] { in });
		if (resp.serviceData.sizeOfPartialErrors() > 0) {
			throw new ServiceException(Arrays.toString(resp.serviceData.getPartialError(0).getMessages()));
		}
		return resp.serviceData.getCreatedObject(0);
	}

	public static TCComponent loadObject(String uid) {
		DataManagementService dmService = DataManagementService.getService(session);
		com.teamcenter.soa.client.model.ServiceData serviceData = dmService.loadObjects(new String[] { uid });
		if ((serviceData != null) && (serviceData.sizeOfPlainObjects() > 0)) {
			TCComponent modelObject = (TCComponent) serviceData.getPlainObject(0);
			dmService.refreshObjects(new TCComponent[] { modelObject });
			return modelObject;
		}
		return null;
	}

	public static void deleteAllJob(TCComponent component) throws TCException {
		component.refresh();
		TCComponent[] allWorkflows = component.getReferenceListProperty("fnd0AllWorkflows");
		if (allWorkflows != null)
			for (TCComponent workflow : allWorkflows)
				workflow.delete();
	}

	public static void syncTcProp(TCProperty sourceTcProp, TCProperty targetTcProp) throws TCException {
		if ((sourceTcProp != null) && (targetTcProp != null)) {
			int sourcePropType = sourceTcProp.getPropertyType();
			int targetPropType = targetTcProp.getPropertyType();

			if (sourcePropType != targetPropType) {
				throw new TCException("源属性与目标属性类型不同，不能同步属性值，请检查!");
			}

			if ((!sourceTcProp.isNotArray()) || (!targetTcProp.isNotArray())) {
				throw new TCException("仅支持单值属性，不支持数组属性，请检查!");
			}

			switch (sourcePropType) {
			case 5:
				targetTcProp.setIntValue(sourceTcProp.getIntValue());
				break;
			case 3:
				targetTcProp.setDoubleValue(sourceTcProp.getDoubleValue());
				break;
			case 8:
				targetTcProp.setStringValue(sourceTcProp.getStringValue());
				break;
			case 6:
				targetTcProp.setLogicalValue(sourceTcProp.getLogicalValue());
				break;
			case 2:
				targetTcProp.setDateValue(sourceTcProp.getDateValue());
				break;
			case 4:
			case 7:
			}
		}
	}

	public static boolean hasJob(TCComponent component) throws Exception {
		if (component != null) {
			TCComponent job = component.getCurrentJob();
			if (job != null) {
				return true;
			}
		}
		return false;
	}

	public static boolean hasPropertyValue(TCComponent cmp, String propName) throws TCException {
		if ((cmp != null) && (cmp.getTCProperty(propName) != null)
				&& (cmp.getTCProperty(propName).getDisplayValue() != null)
				&& (!"".equals(cmp.getTCProperty(propName).getDisplayValue()))) {
			return true;
		}
		return false;
	}

	public static String getPropertyDisplayName(TCComponent cmp, String propertyName) throws TCException {
		if (cmp != null) {
			TCProperty property = cmp.getTCProperty(propertyName);
			if (property != null) {
				return property.getPropertyDisplayName();
			}
		}
		return null;
	}

	public static TCComponent findOrCreateItem(String id, String name, String type) throws TCException {
		TCComponent item = findItem(type, id);
		if (item == null) {
			item = create(type, id, null, name, "", null);
		}
		return item;
	}
}
