package com.dfl.report.common;

import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentReleaseStatus;
import com.teamcenter.rac.kernel.TCComponentReleaseStatusType;
import com.teamcenter.rac.kernel.TCException;

public final class TCComponentReleaseStatusUtils
{
  private static final String RELEASE_STATUS = "ReleaseStatus";
  private static final String RELATION_RELEASE_STATUS_LIST = "release_status_list";

  public static boolean existStatus(InterfaceAIFComponent aif)
    throws TCException
  {
    if ((aif instanceof TCComponent)) {
      TCComponent[] releaseStatusList = ((TCComponent)aif).getReferenceListProperty("release_status_list");
      if ((releaseStatusList != null) && (releaseStatusList.length > 0)) {
        return true;
      }
    }
    return false;
  }

  public static boolean hasStatus(TCComponentItemRevision rev, String status)
    throws TCException
  {
    TCComponent[] releaseStatusList = rev.getReferenceListProperty("release_status_list");
    if (releaseStatusList != null) {
      for (TCComponent com : releaseStatusList) {
        if ((com instanceof TCComponentReleaseStatus)) {
          TCComponentReleaseStatus s = (TCComponentReleaseStatus)com;
          if (status.equals(s.getStringProperty("object_name"))) {
            return true;
          }
        }
      }
    }
    return false;
  }

  public static boolean hasStatus(TCComponentItemRevision rev)
    throws TCException
  {
    TCComponent[] releaseStatusList = rev.getReferenceListProperty("release_status_list");
    if ((releaseStatusList != null) && (releaseStatusList.length > 0)) {
      return true;
    }
    return false;
  }

  public static TCComponent[] getStatus(TCComponent component)
    throws TCException
  {
    return component.getReferenceListProperty("release_status_list");
  }

  public static String getAllStatus(TCComponent cmp)
    throws TCException
  {
    if (cmp != null) {
      TCComponent[] allStatusObjs = getStatus(cmp);
      if (allStatusObjs != null) {
        String finalStatus = "";
        for (int i = 0; i < allStatusObjs.length; i++) {
          if (i == allStatusObjs.length - 1)
            finalStatus = finalStatus + allStatusObjs[i].toDisplayString();
          else {
            finalStatus = finalStatus + allStatusObjs[i].toDisplayString() + ",";
          }
        }
        return finalStatus;
      }
    }
    return "";
  }

  public static TCComponentReleaseStatus getStatusByType(TCComponent component, String statusType)
    throws TCException
  {
    TCComponent[] allStatus = getStatus(component);
    for (TCComponent status : allStatus) {
      if (statusType.equals(status.toInternalString())) {
        return (TCComponentReleaseStatus)status;
      }
    }
    return null;
  }

  public static TCComponentReleaseStatusType getStatusType(String formType) throws TCException {
    return (TCComponentReleaseStatusType)TCUtils.getTypeComponent(formType);
  }

  public static TCComponentReleaseStatus createRelease(String releasedType) throws TCException {
    TCComponentReleaseStatusType statusType = getStatusType("ReleaseStatus");
    TCComponentReleaseStatus statusItem = (TCComponentReleaseStatus)statusType.create(releasedType);
    statusItem.save();
    return statusItem;
  }

  public static void addStatus(TCComponent component, TCComponentReleaseStatus status)
    throws TCException
  {
    if (status != null)
      component.add("release_status_list", status);
  }

  public static void removeAllStatus(TCComponent component)
    throws TCException
  {
    TCComponent[] status = getStatus(component);

    for (TCComponent cmp : status)
      if ((cmp instanceof TCComponentReleaseStatus)) {
        TCComponentReleaseStatus st = (TCComponentReleaseStatus)cmp;
        component.remove("release_status_list", st);
      }
  }

  public static void removeStatus(TCComponent component, String statusName)
    throws TCException
  {
    if (statusName == null) {
      return;
    }

    TCComponent[] status = getStatus(component);

    if (status != null)
      for (TCComponent cmp : status)
        if ((cmp instanceof TCComponentReleaseStatus)) {
          TCComponentReleaseStatus st = (TCComponentReleaseStatus)cmp;

          String name = st.getProperty("object_name");
          if (statusName.equals(name))
            component.remove("release_status_list", st);
        }
  }
}
