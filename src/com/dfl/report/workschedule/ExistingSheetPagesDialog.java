package com.dfl.report.workschedule;

import org.eclipse.swt.widgets.Dialog;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.Table;
import org.eclipse.swt.widgets.TableColumn;
import org.eclipse.swt.widgets.TableItem;
import org.eclipse.swt.widgets.Composite;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;

import org.eclipse.jface.viewers.CellEditor;
import org.eclipse.jface.viewers.CellLabelProvider;
import org.eclipse.jface.viewers.IStructuredContentProvider;
import org.eclipse.jface.viewers.LabelProvider;
import org.eclipse.jface.viewers.TableViewer;
import org.eclipse.jface.viewers.TextCellEditor;
import org.eclipse.jface.viewers.Viewer;
import org.eclipse.jface.viewers.ViewerSorter;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.graphics.Image;
import org.eclipse.wb.swt.SWTResourceManager;
import org.omg.CosCollection.Map;

import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCSession;

import swing2swt.layout.BorderLayout;
import org.eclipse.swt.layout.RowLayout;
import swing2swt.layout.FlowLayout;
import org.eclipse.swt.layout.FormLayout;
import org.eclipse.swt.layout.FormData;
import org.eclipse.swt.layout.FormAttachment;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.graphics.Point;
import org.eclipse.swt.layout.FillLayout;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.widgets.Combo;
import org.eclipse.swt.widgets.Text;

public class ExistingSheetPagesDialog extends Dialog {

	private static class ViewerLabelProvider extends LabelProvider {
		public Image getImage(Object element) {
			return super.getImage(element);
		}

		public String getText(Object element) {
			return super.getText(element);
		}
	}

	private static class Sorter extends ViewerSorter {
		public int compare(Viewer viewer, Object e1, Object e2) {
			Object item1 = e1;
			Object item2 = e2;
			return 0;
		}
	}

	private static class ContentProvider implements IStructuredContentProvider {
		public Object[] getElements(Object inputElement) {
			return new Object[0];
		}

		public void dispose() {
		}

		public void inputChanged(Viewer viewer, Object oldInput, Object newInput) {
		}
	}

	protected Object result;
	protected Shell shell;
//	private Table table;
	private TableViewer tableViewer;
	public LinkedHashMap<String, String> map = new LinkedHashMap<String, String>();
	public ArrayList list = new ArrayList();
	public String sheetname;
	public String nameNO;
	public String model;
	private Label lblFdd;
    private List shlist = new ArrayList();
    private String windowsname;
	/**
	 * Create the dialog.
	 * 
	 * @param parent
	 * @param style
	 */
	public ExistingSheetPagesDialog(Shell parent, int style,List shlist) {
		super(parent, style);
		setText("SWT Dialog");
		this.shlist = shlist;
	}
	public ExistingSheetPagesDialog(Shell parent, int style,List shlist,String windowsname) {
		super(parent, style);
		setText("SWT Dialog");
		this.shlist = shlist;
		this.windowsname=windowsname;
	}

	/**
	 * Open the dialog.
	 * 
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

	protected void centerToScreen(Display display) {
		int nLocationX = display.getClientArea().width / 2 - shell.getSize().x / 2;
		int nLocationY = display.getClientArea().height / 2 - shell.getSize().y / 2;
		shell.setLocation(nLocationX, nLocationY);
	}

	/**
	 * Create contents of the dialog.
	 */
	private void createContents() {
		shell = new Shell(getParent(), getStyle());
		shell.setMinimumSize(new Point(136, 20));
		shell.setImage(SWTResourceManager.getImage(ExistingSheetPagesDialog.class,
				"/com/dfl/report/imags/defaultapplication_16.png"));
		shell.setSize(449, 598);
		if(windowsname!=null && !windowsname.isEmpty()) {
			shell.setText("选择减页sheet");
		}else {
			shell.setText("\u9009\u62E9\u589E\u9875\u4F4D\u7F6E");
		}	
		shell.setLayout(new BorderLayout(0, 0));

		Composite composite = new Composite(shell, SWT.NONE);
		composite.setLayoutData(BorderLayout.NORTH);
		composite.setLayout(null);
		

		lblFdd = new Label(composite, SWT.NONE);
		lblFdd.setForeground(SWTResourceManager.getColor(SWT.COLOR_RED));
		lblFdd.setBounds(152, 9, 267, 17);
		lblFdd.setText("");

		Label lblsheet = new Label(composite, SWT.NONE);
		lblsheet.setBounds(10, 9, 160, 17);
		lblsheet.setText("\u5F53\u524D\u62A5\u8868sheet\u9875\u4FE1\u606F");
		
		Label label = new Label(composite, SWT.NONE);
		label.setForeground(SWTResourceManager.getColor(SWT.COLOR_RED));
		label.setBounds(177, 9, 303, 17);

		Composite composite_1 = new Composite(shell, SWT.NONE);
		composite_1.setLayoutData(BorderLayout.SOUTH);
		composite_1.setLayout(null);

		// 确定按钮
		Button btnNewButton = new Button(composite_1, SWT.NONE);
		btnNewButton.setBounds(251, 5, 71, 27);
		btnNewButton.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				ComformEvent();
			}
		});
		btnNewButton.setText("\u786E\u5B9A");

		// 取消按钮
		Button btnNewButton_1 = new Button(composite_1, SWT.NONE);
		btnNewButton_1.setBounds(328, 5, 72, 27);
		btnNewButton_1.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				shell.dispose();
			}
		});
		btnNewButton_1.setText("\u53D6\u6D88");

		tableViewer = new TableViewer(shell, SWT.BORDER | SWT.CHECK | SWT.FULL_SELECTION | SWT.MULTI);
		Table table = tableViewer.getTable();
		table.setToolTipText("");
		table.setHeaderVisible(true);
		table.setLinesVisible(true);
		table.setLayoutData(BorderLayout.CENTER);

		String tableColumn[] = new String[] { "name", "page" };
		tableViewer.setColumnProperties(tableColumn);// 设置列属性
		TableProvider1 provider_2 = new TableProvider1(tableViewer, tableColumn);
		tableViewer.setContentProvider(provider_2);
		tableViewer.setLabelProvider(provider_2);
		tableViewer.setCellModifier(provider_2);

		TableColumn tblclmnNewColumn = new TableColumn(table, SWT.NONE);
		tblclmnNewColumn.setWidth(334);
		tblclmnNewColumn.setText("Sheet\u540D\u79F0");

		CellEditor acelleditor[] = new CellEditor[2];
		acelleditor[1] = new TextCellEditor(table, 131072);
		tableViewer.setCellEditors(acelleditor);
		
		loadTableData();

	}

	protected void ComformEvent() {
		// TODO Auto-generated method stub
		
		Table table = tableViewer.getTable();
		TableItem[] items = table.getItems();// 获取表格所有的行
		for (int i = 0; i < items.length; i++) {
			TableItem tableItem = items[i];
			if (tableItem.getChecked()) {
				String[] str = new String[2];
				str[0] = tableItem.getText(0);
				list.add(str[0]);
			}
		}
		
		if(windowsname!=null && !windowsname.isEmpty()) {
			if(list.size()<1) {
				lblFdd.setText("请勾选需要减页的sheet！");
				return;
			}
		}else {
			if(list.size()<1) {
				lblFdd.setText("请勾选需要增加的sheet放在哪个sheet页后面！");
				return;
			}
			if(list.size()>1) {
				lblFdd.setText("请勾选单个sheet页！");
				list.clear();
				return;
			}
			sheetname = (String) list.get(0);
		}
		
		shell.dispose();
	}

	private void loadTableData() {
		// TODO Auto-generated method stub
		if(shlist!=null && shlist.size()>0) {
			TableInfo[] infos = new TableInfo[shlist.size()];
			for(int i=0;i<shlist.size();i++) {
				String name = (String) shlist.get(i);
				infos[i] = new TableInfo(name, "", false);
			}
			tableViewer.setInput(infos);
		}
	}

	public static void main(String[] args) {
		List list = new ArrayList();
		list.add("首页");
		list.add("有效页");
		list.add("构成表");
		list.add("构成图");
		ExistingSheetPagesDialog dialog1 = new ExistingSheetPagesDialog(new Shell(), SWT.SHELL_TRIM,list,"选择减页sheet");
		dialog1.open();
	}
}
