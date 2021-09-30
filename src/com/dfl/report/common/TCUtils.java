package com.dfl.report.common;

import com.teamcenter.rac.aif.AbstractAIFOperation;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.IPerspectiveDef;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.common.TCConstants;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentType;
import com.teamcenter.rac.kernel.TCComponentUser;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.ui.common.RACUIUtil;
import com.teamcenter.rac.util.FilterDocument;
import java.util.Date;

public final class TCUtils
{
  public static final String CLIENT_TYPE_IIOP = "IIOP";
  public static final String CLIENT_TYPE_TCCS = "TCCS";
  public static final String CLIENT_TYPE_HTTP = "HTTP";
  public static final String LABEL_MY_TEAMCENTER = "我的 Teamcenter";
  public static final String LABEL_PSE = "结构管理器";
  
  public static String getClientType()
  {
    if (TCSession.isFourTier()) {
      return "HTTP";
    }
    if (TCSession.isTccs()) {
      return "TCCS";
    }
    return "IIOP";
  }

  public static TCSession getTCSession()
  {
    return RACUIUtil.getTCSession();
  }

  public static TCSession getTCSession2()
  {
    return (TCSession)AIFUtility.getDefaultSession();
  }

  public static String getEncoding()
  {
    return TCSession.getServerEncodingName(getTCSession());
  }
  public static void queueOperation(AbstractAIFOperation job)
  {
    getTCSession().queueOperation(job);
  }
  public static int getDefaultMaxNameSize()
  {
    return TCConstants.getDefaultMaxNameSize(getTCSession());
  }
  public static int getDefaultMaxTextAreaSize()
  {
    return 240;
  }

  public static TCPreferenceService getPreferenceService()
  {
    return getTCSession().getPreferenceService();
  }
  public static TCComponentType getTypeComponent(String name)    throws TCException
  {
    return getTCSession().getTypeComponent(name);
  }
  public static FilterDocument createFilterDocument()
  {
    return new FilterDocument(getDefaultMaxNameSize(), getEncoding());
  }

  public static void pasteToNewStuffFloder(TCComponent comp)
    throws TCException
  {
    TCComponentFolder newStuffFolder = getTCSession().getUser().getNewStuffFolder();
    newStuffFolder.add("contents", comp);
  }

  public static Date getTCServerTime()
    throws TCException
  {
    TCComponentFolder folder = getTCSession().getUser().getHomeFolder();
    String desc = folder.getProperty("object_desc");
    folder.setProperty("object_desc", desc + ".");
    Date date = folder.getDateProperty("last_mod_date");
    folder.setProperty("object_desc", desc);
    return date;
  }

  public static AbstractAIFUIApplication getCurrentApplication()
  {
    return AIFUtility.getCurrentApplication();
  }

  public static IPerspectiveDef getCurrentPerspectiveDef()
  {
    return AIFUtility.getCurrentPerspectiveDef();
  }
}
