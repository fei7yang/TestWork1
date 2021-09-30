package com.dfl.report.home;

import org.eclipse.jface.viewers.TreeViewer;
import org.eclipse.swt.SWT;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.widgets.Dialog;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.Tree;
import org.eclipse.swt.widgets.TreeColumn;
import org.eclipse.swt.widgets.TreeItem;
import org.eclipse.wb.swt.SWTResourceManager;

import com.dfl.report.util.Util;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

import swing2swt.layout.BorderLayout;
import swing2swt.layout.FlowLayout;

//import swing2swt.layout.BorderLayout;
//import swing2swt.layout.FlowLayout;

public class OpenHomeDialog extends Dialog {

	protected Object result;
	protected Shell shell;
	private TreeViewer treeViewer;
	
	private Tree tree;
	private TCComponent rootFolder;
	public TCComponent folder;
	public boolean flag=false;
	private TCSession session;

//	private OpenHomeHandler handler;
	
	public OpenHomeDialog(Shell parent,TCComponent rootFolder,TCSession session) {
		super(parent, SWT.SHELL_TRIM);
		this.rootFolder = rootFolder;
		this.session = session;
	}
	
	/**
	 * Create the dialog.
	 * @param parent
	 * @param style
	 */
	public OpenHomeDialog(Shell parent, int style) {
		super(parent, style);
		setText("SWT Dialog");
	}

	/**
	 * Open the dialog.
	 * @return the result
	 */
	public Object open() {
		createContents();
		shell.open();
		shell.layout();
		Display display = getParent().getDisplay();
		centerToScreen(display);
		while (!shell.isDisposed()) {
			if (!display.readAndDispatch()) {
				display.sleep();
			}
		}
		return result;
	}

	
	protected void centerToScreen(Display display)
	{
		int nLocationX=display.getClientArea().width/2-shell.getSize().x/2;
		int nLocationY=display.getClientArea().height/2-shell.getSize().y/2;
		shell.setLocation(nLocationX,nLocationY);
	}
	
	/**
	 * Create contents of the dialog.
	 */
	private void createContents() {
		shell = new Shell(getParent(), getStyle());
		shell.setImage(SWTResourceManager.getImage(OpenHomeDialog.class, "/com/dfl/report/home/productdata_16.png"));
		shell.setSize(450, 538);
		shell.setText("\u9009\u62E9\u6587\u4EF6\u5939");
		
		shell.setLayout(new BorderLayout(0, 0));
		
		Composite composite = new Composite(shell, SWT.NONE);
		composite.setLayoutData(BorderLayout.CENTER);
		composite.setLayout(new BorderLayout(0, 0));
		
		 treeViewer = new TreeViewer(composite, SWT.BORDER);
		 tree = treeViewer.getTree();
		
		tree = treeViewer.getTree();
		tree.setHeaderVisible(true);
		tree.setLinesVisible(true);
		tree.setLayoutData(BorderLayout.CENTER);
		TreeColumn column = new TreeColumn(tree, SWT.NONE);
		column.setImage(SWTResourceManager.getImage(OpenHomeDialog.class, "/com/dfl/report/home/homefolder_16.png"));
		column.setText("Home");
		column.setWidth(420);


		String[] columnProperties1 = new String[]{"name"};
		treeViewer.setColumnProperties(columnProperties1);
		TreeProvider contentProvider = new TreeProvider(treeViewer);
		treeViewer.setContentProvider(contentProvider);
		treeViewer.setLabelProvider(contentProvider);

		treeViewer.setCellModifier(contentProvider);
		
		
		Composite composite_1 = new Composite(shell, SWT.NONE);
		composite_1.setLayoutData(BorderLayout.SOUTH);
		composite_1.setLayout(new FlowLayout(FlowLayout.CENTER, 5, 10));
		
		Button btnNewButton = new Button(composite_1, SWT.NONE);
		btnNewButton.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				okAction();
			}
		});
		btnNewButton.setText("    \u786E\u5B9A    ");
		
		Button btnNewButton_1 = new Button(composite_1, SWT.NONE);
		btnNewButton_1.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				flag = true;
				shell.dispose();
			}
		});
		btnNewButton_1.setText("    \u53D6\u6D88    ");
		
		
		TreeNode rootInput = new TreeNode(rootFolder);
		treeViewer.setInput(rootInput);

	}

	
	
	protected void okAction() {
		// TODO Auto-generated method stub
		System.out.println(">>>ok");
		
		TreeItem[] treeItems = treeViewer.getTree().getSelection();
		if(treeItems==null||treeItems.length<=0)
		{
			return ;
		}
		TreeNode treeNode = (TreeNode) treeItems[0].getData();
		 folder = treeNode.getFolder();
		//handler.doSelectFolder(folder);
		//判断当前用户对所选对象是否有写权限 
		 boolean flag = Util.hasWritePrivilege(session, folder);
		 if(!flag) {
			 folder=null;
		    MessageBox.post("对当前所选文件夹没有写权限！", "温馨提示", MessageBox.INFORMATION);
			return ;
		 }
		 shell.dispose();
		
	}
	
	public static void main(String[] args) {
		OpenHomeDialog dialog = new OpenHomeDialog(new Shell(), SWT.SHELL_TRIM);
		dialog.open();
	}
	
}
