package com.dfl.report.mfcadd;

import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.jface.viewers.ISelectionProvider;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.ui.IViewPart;
import org.eclipse.ui.IViewReference;
import org.eclipse.ui.IWorkbenchPage;
import org.eclipse.ui.handlers.HandlerUtil;

import com.dfl.report.home.OpenHomeDialog;
import com.teamcenter.rac.aif.AIFDesktop;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.cme.accountabilitycheck.Activator;
import com.teamcenter.rac.cme.common.companion.CompanionUtils;
import com.teamcenter.rac.cme.framework.views.primary.BOMPrimaryView;
import com.teamcenter.rac.cme.idc.views.primary.IDCPrimaryView;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentBOMView;
import com.teamcenter.rac.kernel.TCComponentBOMWindow;
import com.teamcenter.rac.kernel.TCComponentBOMWindowType;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.pse.common.BOMLineNode;
import com.teamcenter.rac.pse.common.BOMTreeTable;
import com.teamcenter.rac.util.AdapterUtil;
import com.teamcenter.rac.util.MessageBox;
import com.teamcenter.rac.util.OSGIUtil;
import com.teamcenter.rac.util.PlatformHelper;
import com.teamcenter.rac.vns.model.IContentView;
import com.teamcenter.rac.vns.services.IViewQueryService;

public class VehiclePSummaryReportHandler extends AbstractHandler {
	private AbstractAIFUIApplication app;
	private Shell shell;
	private TCSession session;
	private TCComponentFolder rootFolder;
	private TCComponent savefolder;
	TCComponentBOMLine bopLine = null;
	TCComponentBOMLine ebomLine = null;
	TCComponentBOMWindow ebomWindow = null;

	@Override
	public Object execute(ExecutionEvent event) throws ExecutionException {
		// TODO Auto-generated method stub
		app = AIFUtility.getCurrentApplication();
		session = (TCSession) app.getSession();
		shell = AIFDesktop.getActiveDesktop().getShell();
		InterfaceAIFComponent[] aifComponents = app.getTargetComponents();
		if (aifComponents == null || aifComponents.length != 1) {
			MessageBox.post("������ֻ��ѡ��һ��BBOM�ܳ�B8_BBOMTopNodeRevision����", "����", MessageBox.INFORMATION);
			return null;
		}
		if (aifComponents[0] instanceof TCComponentBOMLine) {
			bopLine = (TCComponentBOMLine) aifComponents[0];
			try {
				if (!bopLine.getItemRevision().isTypeOf("B8_BBOMTopNodeRevision")) {
					MessageBox.post("��ѡ�����д��ڲ���BBOM�ܳ�B8_BBOMTopNodeRevision����", "��ܰ��ʾ", MessageBox.INFORMATION);
					return null;
				}
			} catch (TCException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		} else {
			MessageBox.post("��ѡ�����д��ڲ���BOMLine����", "��ʾ", MessageBox.INFORMATION);
			return null;
		}
		// String designItemid = "";
		// ��ѯ������ģ��
		String inputStream = TemplateUtil.getTemplateFile("DFL_Template_VehicleModleSummary");
		if (inputStream == null) {
			MessageBox.post("����û���ҵ��������ʽ������Ϣ���ܱ��ģ�壬����ϵϵͳ����Ա����TC�����ģ��(����Ϊ��DFL_Template_VehicleModleSummary)��", "��ʾ",
					MessageBox.INFORMATION);
			return null;
		}

		TCComponentItem desginItem = null;
		try {
			List<TCComponent> lstLinks = CompanionUtils.getExistingStructureLinks(bopLine);
			if (lstLinks == null || lstLinks.size() == 0) {
				MFCUtility.errorMassges("��ѡBBOM�ܳ�B8_BBOMTopNodeRevision����δ����EBOM����");
				return null;
			}
			for (int i = 0; i < lstLinks.size(); i++) {
				System.out.println("link is := " + lstLinks.get(i).getType() + " --> " + lstLinks.get(i));
				if (lstLinks.get(i) instanceof TCComponentBOMView) {
					// designItemid =
					// ((TCComponentBOMView)lstLinks.get(i)).getReferenceProperty("parent_item").getProperty("item_id");
					desginItem = (TCComponentItem) ((TCComponentBOMView) lstLinks.get(i))
							.getReferenceProperty("parent_item");
				}
			}
			if (desginItem == null) {
				MFCUtility.errorMassges("��ѡBBOM�ܳ�B8_BBOMTopNodeRevision����δȡ��������EBOM����");
				return null;
			}
			TCComponentBOMWindowType windType = (TCComponentBOMWindowType) session.getTypeComponent("BOMWindow");
			this.ebomWindow = windType.create(null);
			this.ebomLine = this.ebomWindow.setWindowTopLine(desginItem, null, null, null);
//			IViewPart viewpart = (IViewPart)HandlerUtil.getActivePart(event);
//		     IWorkbenchPage localIWorkbenchPage = PlatformHelper.getCurrentPage();
//		      IViewQueryService localIViewQueryService = (IViewQueryService)OSGIUtil.getService(com.teamcenter.rac.cme.accountabilitycheck.Activator.getDefault(), IViewQueryService.class);
//		      List<IViewReference> localList = localIViewQueryService.getAllPrimaryViewReferences(localIWorkbenchPage);
//		      Iterator<IViewReference> itRefs =  localList.iterator();
//		      IViewReference refs = null;
//		      IViewPart cviewpart = null;
//		      while(itRefs.hasNext()) {
//		    	  refs = itRefs.next();
//		    	  cviewpart = (IViewPart) refs.getPart(false);
//		    	  if(cviewpart == null) {
//		    		  continue;
//		    	  }
//		    	  IContentView icv = (IContentView)AdapterUtil.getAdapter(cviewpart, IContentView.class);
//		    	  if(icv != null && cviewpart != viewpart && !icv.isEmpty() && 
//		    			   //(localIWorkbenchPage.isPartVisible(cviewpart)) &&
//		    			   (((cviewpart instanceof BOMPrimaryView)) || ((cviewpart instanceof IDCPrimaryView))))
//		    	  {
//		    		  BOMTreeTable treeTable = (BOMTreeTable)AdapterUtil.getAdapter(cviewpart, BOMTreeTable.class);
//		    		  if(treeTable == null) {
//		    			  continue;
//		    		  }
//		    		  BOMLineNode node = treeTable.getRootBOMLineNode();
//		    		  if (node == null) {
//		    		       continue;
//		    		   }
//		    		  TCComponentBOMLine rootBOMLine = node.getBOMLine();
//		    		  if(rootBOMLine == null) {
//		    			  continue;
//		    		  }
//		    		  try {
//						String type = rootBOMLine.getItemRevision().getType();
//						if(type.equals("DFL9VehicleRevision")  && rootBOMLine.getItem().getProperty("item_id").equals(designItemid)) {
//							ebomLine = rootBOMLine;
//						}
//					} catch (TCException e) {
//						// TODO Auto-generated catch block
//						e.printStackTrace();
//					}
//		    	  }
//		      }
//		      if(ebomLine == null ) {
//		    	  MFCUtility.errorMassges("δչ��BBOM�ܳɹ�����EBOM��ͼ��");
//		    	  return null;
//		      }
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		Thread thread = new Thread() {
			public void run() {
				execute();
			}
		};
		thread.start();
		return null;
	}

	protected void execute() {
		// TODO Auto-generated method stub

		// InterfaceAIFComponent aifComponent = app.getTargetComponent();

		try {
			rootFolder = session.getUser().getHomeFolder();
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		// rootFolder = (TCComponent) aifComponent;

		Display.getDefault().asyncExec(new Runnable() {
			@Override
			public void run() {
				openDialog();
			}
		});

	}

	protected void openDialog() {
		// TODO Auto-generated method stub
		OpenHomeDialog dialog = new OpenHomeDialog(shell, rootFolder, session);
		dialog.open();

		savefolder = dialog.folder;
		System.out.println("�ļ��У�" + dialog.folder);

		if (dialog.flag) {
			return;
		}

		if (savefolder == null) {
			return;
		}
		VehiclePSummaryReportAction action = new VehiclePSummaryReportAction(this.bopLine, ebomLine, ebomWindow,
				savefolder);
		new Thread(action).start();
	}
}
