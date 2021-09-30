package com.dfl.report.common;

import com.teamcenter.rac.kernel.ServiceData;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentGroupMember;
import com.teamcenter.rac.kernel.TCComponentProject;
import com.teamcenter.rac.kernel.TCComponentProjectType;
import com.teamcenter.rac.kernel.TCComponentUser;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.schemas.soa._2006_03.exceptions.ServiceException;
import com.teamcenter.services.rac.core.ProjectLevelSecurityService;
import com.teamcenter.services.rac.core._2012_09.ProjectLevelSecurity;

import java.io.PrintStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public final class TCComponentProjectUtils
{
  private static final String TYPE_TC_PROJECT = "TC_Project";
  private static final String PROP_PROJECT_LIST = "project_list";
  private static final String PROP_PROJECT_ID = "project_id";
  private static final String PROP_PROJECT_NAME = "project_name";
  private static TCSession sesssion = TCUtils.getTCSession();

//  public static TCComponentProject createProject(String projectId, String projectName, String projectDesc, TCComponentGroupMember[] teamMembers, TCComponentUser[] privileges, TCComponentUser[] teamAdministrators)
//    throws ServiceException
//  {
//    List teamMemberList = new ArrayList();
//
//    if (teamMembers != null) {
//      for (TCComponentGroupMember member : teamMembers)
//      {
//        ProjectLevelSecurity.TeamMemberInfo info = new ProjectLevelSecurity.TeamMemberInfo();
//        info.teamMemberType = 0;
//        info.teamMember = member;
//        teamMemberList.add(info);
//      }
//    }
//
//    if (privileges != null) {
//      for (TCComponentUser privilege : privileges)
//      {
//        ProjectLevelSecurity.TeamMemberInfo info = new ProjectLevelSecurity.TeamMemberInfo();
//        info.teamMemberType = 1;
//        info.teamMember = privilege;
//        teamMemberList.add(info);
//      }
//    }
//
//    if (teamAdministrators != null) {
//      for (TCComponentUser teamAdmin : teamAdministrators)
//      {
//        ProjectLevelSecurity.TeamMemberInfo info = new ProjectLevelSecurity.TeamMemberInfo();
//        info.teamMemberType = 2;
//        info.teamMember = teamAdmin;
//        teamMemberList.add(info);
//      }
//    }
//
//    ProjectLevelSecurity.TeamMemberInfo[] allTeamMembers = new ProjectLevelSecurity.TeamMemberInfo[teamMemberList.size()];
//    for (int i = 0; i < teamMemberList.size(); i++) {
//      ProjectLevelSecurity.TeamMemberInfo member = (ProjectLevelSecurity.TeamMemberInfo)teamMemberList.get(i);
//      allTeamMembers[i] = member;
//    }
//
//    ProjectLevelSecurityService projectService = ProjectLevelSecurityService.getService(sesssion);
//    ProjectLevelSecurity.ProjectInformation[] projectInfos = new ProjectLevelSecurity.ProjectInformation[1];
//    ProjectLevelSecurity.ProjectInformation projectInfo = new ProjectLevelSecurity.ProjectInformation();
//    projectInfo.active = true;
//    projectInfo.clientId = "PLS-RAC-SESSION";
//    projectInfo.projectId = projectId;
//    projectInfo.projectName = projectName;
//    projectInfo.projectDescription = projectDesc;
//    projectInfo.useProgramContext = false;
//    projectInfo.visible = true;
//    projectInfo.teamMembers = allTeamMembers;
//    projectInfos[0] = projectInfo;
//    ProjectLevelSecurity.ProjectOpsResponse response = projectService.createProjects(projectInfos);
//    TCServiceExceptionUtils.validateServiceException(response.serviceData);
//    if (response.serviceData.sizeOfCreatedObjects() > 0) {
//      return (TCComponentProject)response.serviceData.getCreatedObject(0);
//    }
//    return null;
//  }

  public static void assignGroupMemberToProject(TCComponentProject project, TCComponent[] groupMembers, TCComponentUser[] privilegedUsers, TCComponentUser[] teamAdministrators)
    throws TCException
  {
    if ((project != null) && (groupMembers != null))
      project.modifyTeam(groupMembers, teamAdministrators, privilegedUsers);
  }

//  public static List<TCComponentProject> findProjectsByUserId(String userId)
//    throws TCException
//  {
//    TCComponent[] comps = TCComponentQueryUtils.query("基于用户的项目", "ID", userId);
//    if ((comps != null) && (comps.length > 0)) {
//      List projectList = new ArrayList();
//      for (TCComponent comp : comps) {
//        if ((comp instanceof TCComponentProject)) {
//          projectList.add((TCComponentProject)comp);
//        }
//      }
//      return projectList;
//    }
//    return null;
//  }

//  public static List<String> findProjectIdListByUserId(String userId)
//    throws TCException
//  {
//    List projectList = findProjectsByUserId(userId);
//    List projectIdList = new ArrayList();
//    if (projectList != null) {
//      for (TCComponentProject project : projectList) {
//        projectIdList.add(project.getStringProperty("project_id") + "-" + 
//          project.getStringProperty("project_name"));
//      }
//    }
//    return projectIdList;
//  }

  public static TCComponentProject findProject(String projectId)
    throws TCException
  {
    TCComponentProjectType type = (TCComponentProjectType)sesssion.getTypeComponent("TC_Project");
    return type.find(projectId);
  }

//  public static TCComponentProject findProjectByName(String projectName)
//    throws TCException
//  {
//    TCComponent[] comps = TCComponentQueryUtils.queryByProperty(Messages.QueryUtil_PROJECTS, "project_name", 
//      projectName);
//    if (comps != null) {
//      if (comps.length == 0)
//        return null;
//      if (comps.length == 1) {
//        return (TCComponentProject)comps[0];
//      }
//      throw new TCException("项目[" + projectName + "]存在多个!");
//    }
//
//    return null;
//  }

//  public static TCComponentProject[] findProjectsByName(String projectName)
//    throws TCException
//  {
//    TCComponent[] cmps = TCComponentQueryUtils.queryByProperty(Messages.QueryUtil_PROJECTS, "project_name", 
//      projectName);
//    if ((cmps != null) && (cmps.length > 0)) {
//      TCComponentProject[] projects = new TCComponentProject[cmps.length];
//      for (int i = 0; i < cmps.length; i++) {
//        projects[i] = ((TCComponentProject)cmps[i]);
//      }
//      return projects;
//    }
//    return null;
//  }
//
//  public static List<TCComponentProject> findAllProjects()
//    throws TCException
//  {
//    TCComponent[] comps = TCComponentQueryUtils.query(Messages.QueryUtil_PROJECTS, Messages.QueryUtil_PROJECTID, 
//      "*");
//    if ((comps != null) && (comps.length > 0)) {
//      List projectList = new ArrayList();
//      TCComponent[] arrayOfTCComponent1 = comps; int j = comps.length; for (int i = 0; i < j; i++) { TCComponent comp = arrayOfTCComponent1[i];
//        if ((comp instanceof TCComponentProject)) {
//          projectList.add((TCComponentProject)comp);
//        }
//      }
//      return projectList;
//    }
//    return null;
//  }

  public static void assignProject(TCComponent component, TCComponentProject project)
    throws TCException
  {
    project.assignToProject(new TCComponent[] { component });
  }

  public static TCComponent[] getTCComponentProjects(TCComponent component)
    throws TCException
  {
    return TCComponentUtils.getCompsByRelation(component, "project_list");
  }

  public static void removeProject(TCComponent component, TCComponentProject project)
    throws TCException
  {
    project.removeFromProject(new TCComponent[] { component });
  }

  public static void assignProject(List<TCComponent> componentList, TCComponentProject project)
    throws TCException
  {
    TCComponent[] components = new TCComponent[componentList.size()];
    for (int i = 0; i < componentList.size(); i++) {
      components[i] = ((TCComponent)componentList.get(i));
    }
    project.assignToProject(components);
  }

  public static void removeProject(List<TCComponent> componentList, TCComponentProject project)
    throws TCException
  {
    TCComponent[] components = new TCComponent[componentList.size()];
    for (int i = 0; i < componentList.size(); i++) {
      components[i] = ((TCComponent)componentList.get(i));
    }
    project.removeFromProject(components);
  }

  public static TCComponentProject getFirstProject(TCComponent component)
    throws TCException
  {
    if (component != null) {
      TCComponent[] projectList = component.getReferenceListProperty("project_list");
      if ((projectList != null) && (projectList.length > 0)) {
        return (TCComponentProject)projectList[0];
      }
    }
    return null;
  }

  public static String getFisrtProjectId(TCComponent component)
    throws TCException
  {
    TCComponentProject firstProject = getFirstProject(component);
    if (firstProject != null) {
      return firstProject.getStringProperty("project_id");
    }
    return null;
  }

  public static String getFisrtProjectName(TCComponent component)
    throws TCException
  {
    TCComponentProject firstProject = getFirstProject(component);
    if (firstProject != null) {
      return firstProject.getStringProperty("project_name");
    }
    return null;
  }

//  public static List<TCComponentGroupMember> getGroupMemberList(TCComponentProject project)
//    throws TCException
//  {
//    if (project == null) {
//      throw new RuntimeException("项目不存在，请检查!");
//    }
//
//    List teams = project.getTeam();
//    if (teams != null) {
//      List groupMemberList = new ArrayList();
//      int j;
//      int i;
//      label119: for (Iterator localIterator = teams.iterator(); localIterator.hasNext(); 
//        i < j)
//      {
//        Object obj = localIterator.next();
//        TCComponent[] cmps = (TCComponent[])obj;
//        if ((cmps == null) || (cmps.length <= 0))
//          break label119;
//        TCComponent[] arrayOfTCComponent1;
//        j = (arrayOfTCComponent1 = cmps).length; i = 0; continue; TCComponent cmp = arrayOfTCComponent1[i];
//        if ((cmp instanceof TCComponentGroupMember))
//          groupMemberList.add((TCComponentGroupMember)cmp);
//        i++;
//      }
//
//      return groupMemberList;
//    }
//    return null;
//  }

  public static String getProjectOwner(TCComponentProject project)
    throws Exception
  {
    if (project != null) {
      TCComponent owningUser = project.getReferenceProperty("owning_user");
      if (owningUser != null) {
        return owningUser.getStringProperty("user_name");
      }
    }
    return "";
  }

//  public static String getProjectManager(TCComponentProject project)
//    throws Exception
//  {
//    String projectManager = "";
//
//    if (project == null) {
//      throw new RuntimeException("项目不存在，请检查!");
//    }
//
//    List teams = project.getTeam();
//    if (teams != null)
//    {
//      int j;
//      int i;
//      label179: for (Iterator localIterator = teams.iterator(); localIterator.hasNext(); 
//        i < j)
//      {
//        Object obj = localIterator.next();
//        TCComponent[] cmps = (TCComponent[])obj;
//        if ((cmps == null) || (cmps.length <= 0))
//          break label179;
//        TCComponent[] arrayOfTCComponent1;
//        j = (arrayOfTCComponent1 = cmps).length; i = 0; continue; TCComponent cmp = arrayOfTCComponent1[i];
//        if ((cmp instanceof TCComponentGroupMember)) {
//          TCComponentGroupMember groupMember = (TCComponentGroupMember)cmp;
//          System.out.println(groupMember.toDisplayString());
//        }
//        if ((cmp instanceof TCComponentUser)) {
//          TCComponentUser user = (TCComponentUser)cmp;
//          System.out.println(user.toDisplayString());
//          projectManager = projectManager + user.getStringProperty("user_name") + ",";
//        }
//        i++;
//      }
//
//    }
//
//    if ("".equals(projectManager)) {
//      throw new RuntimeException("项目[" + project.toDisplayString() + "]中不存在项目经理，请检查项目管理员角色!");
//    }
//
//    projectManager = projectManager.substring(0, projectManager.length() - 1);
//
//    return projectManager;
//  }

//  public static TCComponentUser getUserInProject(TCComponentProject project, String userName)
//    throws TCException
//  {
//    if (project == null) {
//      throw new RuntimeException("项目不存在，请检查!");
//    }
//
//    if ((userName == null) || ("".equals(userName))) {
//      throw new RuntimeException("用户名不能为空，请检查!");
//    }
//
//    List teams = project.getTeam();
//    if (teams != null)
//    {
//      int j;
//      int i;
//      label194: for (Iterator localIterator = teams.iterator(); localIterator.hasNext(); 
//        i < j)
//      {
//        Object obj = localIterator.next();
//        TCComponent[] cmps = (TCComponent[])obj;
//        if ((cmps == null) || (cmps.length <= 0))
//          break label194;
//        TCComponent[] arrayOfTCComponent1;
//        j = (arrayOfTCComponent1 = cmps).length; i = 0; continue; TCComponent cmp = arrayOfTCComponent1[i];
//        if ((cmp instanceof TCComponentGroupMember)) {
//          TCComponentGroupMember groupMember = (TCComponentGroupMember)cmp;
//          TCComponentUser groupMemberUser = groupMember.getUser();
//          if ((groupMemberUser != null) && 
//            (userName.equals(groupMemberUser.getStringProperty("user_name")))) {
//            return groupMemberUser;
//          }
//        }
//        if ((cmp instanceof TCComponentUser)) {
//          TCComponentUser priUser = (TCComponentUser)cmp;
//          if (userName.equals(priUser.getStringProperty("user_name")))
//            return priUser;
//        }
//        i++;
//      }
//
//    }
//
//    return null;
//  }
}
