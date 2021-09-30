package com.dfl.report.common;

import com.teamcenter.rac.aif.AIFClipboard;
import com.teamcenter.rac.aif.AIFPortal;
import com.teamcenter.rac.aif.AIFTransferable;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.services.ISelectionMediatorService;
import com.teamcenter.rac.util.MessageBox;
import java.awt.Toolkit;
import java.awt.Window;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;
import java.io.PrintStream;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.util.ArrayList;
import java.util.List;
import javax.swing.JOptionPane;
import javax.swing.SwingUtilities;
import org.osgi.framework.Bundle;
import org.osgi.framework.BundleContext;
import org.osgi.framework.FrameworkUtil;

public final class EclipseUtils
{
  public static final int ERROR = 1;
  public static final int INFOMATION = 2;
  public static final int WARNING = 4;
  
  public static void invokeLater(Runnable runnable)
  {
    SwingUtilities.invokeLater(runnable);
  }
  public static MessageBox error(Throwable throwable)
  {
    StringWriter stringWriter = new StringWriter();
    PrintWriter writer = new PrintWriter(stringWriter);
    throwable.printStackTrace(writer);
    
    return MessageBox.post(stringWriter.toString(), "Exception", 1);
  }

  public static MessageBox info(String msg)
  {
    return MessageBox.post(msg, Messages.INFOMATION, 2);
  }

  public static MessageBox warn(String msg)
  {
    return MessageBox.post(msg, Messages.WARNING, 4);
  }

  public static MessageBox error(String msg)
  {
    return MessageBox.post(msg, Messages.ERROR, 1);
  }
  public static MessageBox showMsg(String title, String msg, int type)
  {
    return MessageBox.post(msg, title, type);
  }
  public static MessageBox showMsg(Window window, String title, String msg, int type)
  {
    return MessageBox.post(window, msg, title, type);
  }

  public static boolean confirm(String title, String message)
  {
    int num = JOptionPane.showConfirmDialog(
    
      null, message, title, 0);
    return num == 0;
  }

  public static InterfaceAIFComponent[] getTargetComponents()
  {
    BundleContext bundleCtx = FrameworkUtil.getBundle(EclipseUtils.class).getBundleContext();
    ISelectionMediatorService selectionSvc = 
      (ISelectionMediatorService)bundleCtx.getService(bundleCtx.getServiceReference(ISelectionMediatorService.class.getName()));
    InterfaceAIFComponent[] pasteTargets = selectionSvc.getTargetComponents();
    
    return pasteTargets;
  }

  public static InterfaceAIFComponent getTargetComponent()
  {
    InterfaceAIFComponent[] targets = getTargetComponents();
    if ((targets != null) && (targets.length >= 1)) {
      return targets[0];

    }
    return null;
  }

  public static List<InterfaceAIFComponent> getComponentsFromClipboard()
  {
    AIFClipboard clipboard = AIFPortal.getClipboard();
    if (clipboard != null)    {
      Transferable contents = clipboard.getContents(null);
      if (contents != null) {
        return ((AIFTransferable)contents).getComponents();
      }
    }
    return null;
  }


  public static String getTcClipboard()
  {
    return null;
  }


//  public static List<TCComponent> getTcObjFormClipboard()
//  {
//    String contents = getSystemClipboardStr();
//    if (contents != null)    {
//      List<TCComponent> list = new ArrayList();
//      String[] contentArray = contents.split("\r\n");
//      if (contentArray != null)        {     
//        String[] arrayOfString1;
//        int j = (arrayOfString1 = contentArray).length;  
//        for (int i = 0; i < j; i++)    {         
//        String contentStr = arrayOfString1[i];
//          int oIndex = 0;
//          int sidIndex = 0;
//          String clientType = TCUtils.getClientType();
//          
//          String objKey = "&-o=";
//          if (contentStr.contains(objKey)) {
//            oIndex = contentStr.indexOf(objKey) + objKey.length();
//
//
//          }
//          String sidKey = "&-sid=";
//          if (("IIOP".equals(clientType)) && (contentStr.contains(sidKey))) {
//            sidIndex = contentStr.indexOf(sidKey);
//
//
//          }
//          String serverNameKey = "&servername=";
//          if (("HTTP".equals(clientType)) && (contentStr.contains(serverNameKey))) {
//            sidIndex = contentStr.indexOf(serverNameKey);
//
//          }
//          if ((oIndex != 0) && (sidIndex != 0))          {
//            String uid = contentStr.substring(oIndex, sidIndex).replace("AAAAAAAAAAAAA", "");
//            System.out.println(uid);
//            TCComponent component = TCComponentUtils.loadObject(uid);
//            if (component != null) {
//              list.add(component);
//            }
//          }
//        }
//      }
//      return list;
//    }
//    return null;
//  }

  public static String getSystemClipboardStr()
  {
    String ret = "";
    Clipboard sysClip = Toolkit.getDefaultToolkit().getSystemClipboard();
    

    Transferable clipTf = sysClip.getContents(null);

    if (clipTf != null) {
      try
      {
        ret = (String)clipTf.getTransferData(DataFlavor.stringFlavor);
      }      catch (Exception e)      {
        e.printStackTrace();

      }
    }
    return ret;
  }
}
