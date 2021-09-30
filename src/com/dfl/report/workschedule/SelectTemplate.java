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

public class SelectTemplate extends Dialog {

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
	private Combo combo_1;
	public LinkedHashMap<String, String> map = new LinkedHashMap<String, String>();
	public ArrayList list = new ArrayList();
	public String nameNO;
	public String model;
	public boolean IsSameout;
	private Label lblNewLabel;
	private Text text;
	private Combo combo;

	/**
	 * Create the dialog.
	 * 
	 * @param parent
	 * @param style
	 */
	public SelectTemplate(Shell parent, int style) {
		super(parent, style);
		setText("SWT Dialog");
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
		//shell.setImage(SWTResourceManager.getImage(SelectTemplate.class,"/com/dfl/report/workschedule/defaultapplication_16.png"));
		shell.setImage(SWTResourceManager.getImage(SelectTemplate.class,"/com/dfl/report/imags/defaultapplication_16.png"));
		shell.setSize(600, 621);
		shell.setText("\u9009\u62E9\u6A21\u677F");
		shell.setLayout(new BorderLayout(0, 0));

		Composite composite = new Composite(shell, SWT.NONE);
		composite.setLayoutData(BorderLayout.NORTH);
		composite.setLayout(null);

		Label label = new Label(composite, SWT.NONE);
		label.setBounds(211, 9, 67, 17);
		label.setText("\u547D\u540D\u5E8F\u53F7\uFF1A");

		Label label_1 = new Label(composite, SWT.NONE);
		label_1.setBounds(10, 9, 61, 17);
		label_1.setText("\u6A21\u677F\u7C7B\u578B\uFF1A");

		combo_1 = new Combo(composite, SWT.NONE);
		combo_1.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				SelectloadTableData();
			}
		});
		combo_1.setBounds(72, 6, 120, 20);
		combo_1.add("��ͨ��λģ��");
		combo_1.add("VIN����ģ��");
		combo_1.add("������ģ��");
		combo_1.select(0);
		
		text = new Text(composite, SWT.BORDER);
		text.setBounds(284, 6, 88, 23);
		
		Label label_2 = new Label(composite, SWT.NONE);
		label_2.setBounds(389, 9, 88, 17);
		label_2.setText("\u5DE6\u53F3\u5DE5\u4F4D\u540C\u51FA\uFF1A");
		
		combo = new Combo(composite, SWT.NONE);
		combo.setBounds(483, 6, 75, 25);
		combo.add("��");
		combo.add("��");
		combo.select(0);

		Composite composite_1 = new Composite(shell, SWT.NONE);
		composite_1.setLayoutData(BorderLayout.SOUTH);
		composite_1.setLayout(null);

		// ȷ����ť
		Button btnNewButton = new Button(composite_1, SWT.NONE);
		btnNewButton.setBounds(388, 5, 71, 27);
		btnNewButton.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				ComformEvent();
			}
		});
		btnNewButton.setText("\u786E\u5B9A");

		// ȡ����ť
		Button btnNewButton_1 = new Button(composite_1, SWT.NONE);
		btnNewButton_1.setBounds(465, 5, 72, 27);
		btnNewButton_1.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				shell.dispose();
			}
		});
		btnNewButton_1.setText("\u53D6\u6D88");
		
		lblNewLabel = new Label(composite_1, SWT.NONE);
		lblNewLabel.setForeground(SWTResourceManager.getColor(SWT.COLOR_RED));
		lblNewLabel.setBounds(10, 10, 372, 17);

		Composite composite_2 = new Composite(shell, SWT.NONE);
		composite_2.setLayoutData(BorderLayout.EAST);
		composite_2.setLayout(new RowLayout(SWT.VERTICAL));

		Button moveUpButton = new Button(composite_2, SWT.NONE);
		moveUpButton.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				moveUpAction();
			}
		});
		moveUpButton.setImage(
				SWTResourceManager.getImage(SelectTemplate.class, "/com/dfl/report/imags/moveup_16.png"));

		Button moveDownButton = new Button(composite_2, SWT.NONE);
		moveDownButton.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				moveDownAction();
			}
		});
		moveDownButton.setImage(
				SWTResourceManager.getImage(SelectTemplate.class, "/com/dfl/report/imags/movedown_16.png"));

		// ȫѡ�¼�
		Button button = new Button(composite_2, SWT.NONE);
		button.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				AllSelectionOp();
			}
		});
		button.setText("\u5168\u90E8\u9009\u62E9");

		// ȡ���¼�
		Button button_1 = new Button(composite_2, SWT.NONE);
		button_1.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				AllCancelOp();
			}
		});
		button_1.setText("\u5168\u90E8\u53D6\u6D88");

		tableViewer = new TableViewer(shell, SWT.BORDER | SWT.CHECK | SWT.FULL_SELECTION | SWT.MULTI);
		Table table = tableViewer.getTable();
		table.setHeaderVisible(true);
		table.setLinesVisible(true);
		table.setLayoutData(BorderLayout.CENTER);

		String tableColumn[] = new String[] { "name", "page" };
		tableViewer.setColumnProperties(tableColumn);// ����������
		TableProvider1 provider_2 = new TableProvider1(tableViewer, tableColumn);
		tableViewer.setContentProvider(provider_2);
		tableViewer.setLabelProvider(provider_2);
		tableViewer.setCellModifier(provider_2);

		TableColumn tblclmnNewColumn = new TableColumn(table, SWT.NONE);
		tblclmnNewColumn.setWidth(222);
		tblclmnNewColumn.setText("Sheet\u540D\u79F0");

		TableColumn tblclmnPage = new TableColumn(table, SWT.NONE);
		tblclmnPage.setWidth(100);
		tblclmnPage.setText("\u9875\u6570");

		CellEditor acelleditor[] = new CellEditor[2];
		acelleditor[1] = new TextCellEditor(table, 131072);
		tableViewer.setCellEditors(acelleditor);
		
//		ArrayList list = getSizeRule();
//		if(list!=null && list.size()>0) {
//			for(int i=0;i<list.size();i++) {
//				combo.add((String) list.get(i));
//			}
//		}else {
//			lblNewLabel.setText("�����Ϣ��ѡ��(DFL9_get_version_information)δ���ã�");
//		}

		loadTableData();

	}

	protected void ComformEvent() {
		// TODO Auto-generated method stub
		if (combo_1.getSelectionIndex() == -1) {
			lblNewLabel.setText("��ѡ��ģ�����ͣ�");
			return;
		}
		if (text.getText().isEmpty()) {
			lblNewLabel.setText("����д������ţ�");
			return;
		}
		if (combo.getText().isEmpty()) {
			lblNewLabel.setText("��ѡ�����ҹ�λ�Ƿ�ͬ����");
			return;
		}
		if(combo.getText().equals("��")) {
			IsSameout = true;
		}else {
			IsSameout = false;
		}		
		model = combo_1.getItem(combo_1.getSelectionIndex());
		nameNO = text.getText();
		Table table = tableViewer.getTable();
		TableItem[] items = table.getItems();// ��ȡ������е���
		String rswqd = "";
		String rswsf = "";
		for (int i = 0; i < items.length; i++) {
			TableItem tableItem = items[i];
			if (tableItem.getChecked()) {
				String[] str = new String[2];
				str[0] = tableItem.getText(0);
				str[1] = tableItem.getText(1);
				if(str[0].equals("RSW����")) {
					rswqd = str[0];
				}
				if(str[0].equals("RSW�ŷ�")) {
					rswsf = str[0];
				}
				if(str[1]==null || str[1].isEmpty()) {
					str[1] = "1";
				}
				map.put(str[0], str[1]);
				list.add(str[0]);
			}
		}
		if(map.size()<1) {
			lblNewLabel.setText("�빴ѡ��Ҫ�����sheetģ�壡");
			return;
		}
		if(!rswqd.isEmpty()&&!rswsf.isEmpty()) {
			lblNewLabel.setText("RSW������RSW�ŷ�ֻ��ѡһ�֣�");
			rswqd = "";
			rswsf = "";
			list.clear();
			for(Iterator<java.util.Map.Entry<String, String>> it = map.entrySet().iterator();it.hasNext();) {
				java.util.Map.Entry<String, String> item = it.next();
				it.remove();
			}

			return;
		}
		
		shell.dispose();
	}

	// ȫѡ����
	protected void AllSelectionOp() {
		Table table = tableViewer.getTable();
		TableItem[] items = table.getItems();// ��ȡ������е���
		ArrayList list = new ArrayList();
		for (int i = 0; i < items.length; i++) {
			TableItem tableItem = items[i];
			tableItem.setChecked(true);
			// tableItem.getChecked();
		}
	}

	// ȫ��ȡ������
	protected void AllCancelOp() {
		Table table = tableViewer.getTable();
		TableItem[] items = table.getItems();// ��ȡ������е���
		ArrayList list = new ArrayList();
		for (int i = 0; i < items.length; i++) {
			TableItem tableItem = items[i];
			tableItem.setChecked(false);
			// tableItem.getChecked();
		}
	}

	protected void moveDownAction() {
		// TODO Auto-generated method stub
		Table table = tableViewer.getTable();

		int[] selectIndices = table.getSelectionIndices();

		if (selectIndices == null || selectIndices.length <= 0) {
			return;
		}

		int[] selectIndices2 = new int[selectIndices.length];

		// ѡ��Ĳ�����������,������
		{
			int tempv = 0;
			for (int i = 0; i < selectIndices.length; i++) {

				selectIndices2[i] = selectIndices[i] + 1;
				if (i == 0) {
					tempv = selectIndices[i];
				} else {
					int t = selectIndices[i] - tempv;
					if (t != 1) {
						System.out.println("ѡ��Ĳ����������������");
						return;
					}
					tempv = selectIndices[i];
				}

			}
		}

		int startRow = selectIndices[0];
		int maxRow = selectIndices[selectIndices.length - 1];

		// ѡ������Ѿ������,������
		{
			if (maxRow == table.getItemCount() - 1) {
				System.out.println("ѡ������Ѿ������,������");
				return;
			}
		}

		// ��ѡ����������ƶ�
		{

			TableItem[] items = table.getItems();// ��ȡ������е���
			ArrayList list = new ArrayList();
			java.util.Map<TableInfo,Boolean> map = new HashMap<>();
			for (int i = 0; i < items.length; i++) {
				TableItem tableItem = items[i];
				// tableItem.getChecked();				
				TableInfo info = (TableInfo) tableItem.getData();
				map.put(info, tableItem.getChecked());
				list.add(info);
			}
			System.out.println("startRow:" + startRow);
			System.out.println("maxRow:" + maxRow);

			ArrayList infoList = MoveUtil.moveDown(list, startRow, maxRow);

			for (int i = 0; i < infoList.size(); i++) {
				TableInfo info = (TableInfo) infoList.get(i);
				System.out.println("info:" + info.getName());
			}

			TableInfo[] newInfos = (TableInfo[]) infoList.toArray(new TableInfo[infoList.size()]);

			table.removeAll();
			tableViewer.setInput(newInfos);
			
			TableItem[] item2 = table.getItems();// ��ȡ������е���
			for (int i = 0; i < item2.length; i++) {
				if(map.containsKey(item2[i].getData()))
				{
					item2[i].setChecked(map.get(item2[i].getData()));
				}
			}

			// ѡ���ƶ������
			table.setSelection(selectIndices2);

		}

	}

	protected void moveUpAction() {
		// TODO Auto-generated method stub
		Table table = tableViewer.getTable();

		int[] selectIndices = table.getSelectionIndices();

		if (selectIndices == null || selectIndices.length <= 0) {
			return;
		}

		int[] selectIndices2 = new int[selectIndices.length];

		// ѡ��Ĳ�����������,������
		{
			int tempv = 0;
			for (int i = 0; i < selectIndices.length; i++) {

				selectIndices2[i] = selectIndices[i] - 1;
				if (i == 0) {
					tempv = selectIndices[i];
				} else {
					int t = selectIndices[i] - tempv;
					if (t != 1) {
						System.out.println("ѡ��Ĳ����������������");
						return;
					}
					tempv = selectIndices[i];
				}

			}
		}

		int startRow = selectIndices[0];
		int maxRow = selectIndices[selectIndices.length - 1];

		// ѡ������Ѿ��ǵ�һ��,������
		{
			if (maxRow == 0) {
				System.out.println("ѡ������Ѿ��ǵ�һ��,������");
				return;
			}
		}

		// ��ѡ����������ƶ�
		{

			TableItem[] items = table.getItems();// ��ȡ������е���
			ArrayList list = new ArrayList();
			java.util.Map<TableInfo,Boolean> map = new HashMap<>();
			for (int i = 0; i < items.length; i++) {
				TableItem tableItem = items[i];
				// tableItem.getChecked();
				TableInfo info = (TableInfo) tableItem.getData();
				map.put(info, tableItem.getChecked());
				list.add(info);
			}
			System.out.println("startRow:" + startRow);
			System.out.println("maxRow:" + maxRow);

			ArrayList infoList = MoveUtil.moveUp(list, startRow, maxRow);

			for (int i = 0; i < infoList.size(); i++) {
				TableInfo info = (TableInfo) infoList.get(i);
				System.out.println("info:" + info.getName());
			}

			TableInfo[] newInfos = (TableInfo[]) infoList.toArray(new TableInfo[infoList.size()]);

			table.removeAll();
			tableViewer.setInput(newInfos);
			
			TableItem[] item2 = table.getItems();// ��ȡ������е���
			for (int i = 0; i < item2.length; i++) {
				if(map.containsKey(item2[i].getData()))
				{
					item2[i].setChecked(map.get(item2[i].getData()));
				}
			}

			// ѡ���ƶ������
			table.setSelection(selectIndices2);

		}

	}

	private void loadTableData() {
		// TODO Auto-generated method stub
		TableInfo[] infos = new TableInfo[25];
		infos[0] = new TableInfo("��ҳ", "", false);
		infos[1] = new TableInfo("��Чҳ", "", false);
		infos[2] = new TableInfo("���ɱ�", "", false);
		infos[3] = new TableInfo("����ͼ", "", false);
		infos[4] = new TableInfo("ʽ����", "", false);
		infos[5] = new TableInfo("Ϳ��", "1", true);
		infos[6] = new TableInfo("������", "1", true);
		infos[7] = new TableInfo("��ĸ��", "1", true);
		infos[8] = new TableInfo("��˨��", "1", true);
		infos[9] = new TableInfo("�㺸-PSW", "1", true);
		infos[10] = new TableInfo("�㺸-RSW", "1", true);
		infos[11] = new TableInfo("�㺸-MSW", "1", true);
		infos[12] = new TableInfo("�㺸-SSW", "1", true);
		infos[13] = new TableInfo(" PSW", "", false);
		infos[14] = new TableInfo("RSW����", "", false);
		infos[15] = new TableInfo("RSW�ŷ�", "", false);
		infos[16] = new TableInfo("��ʱ����", "1", true);
		infos[17] = new TableInfo("������ҵ", "1", true);
		infos[18] = new TableInfo("HEMָʾ", "1", true);
		infos[19] = new TableInfo("������װ", "1", true);
		infos[20] = new TableInfo("װ��", "1", true);
		infos[21] = new TableInfo("���ͳ�Ʊ�", "", false);
		infos[22] = new TableInfo("NEPID��װ", "1", true);
		infos[23] = new TableInfo("����", "1", true);
		infos[24] = new TableInfo("�����߼��", "1", true);
		tableViewer.setInput(infos);

	}

	private void SelectloadTableData() {

		model = combo_1.getItem(combo_1.getSelectionIndex());
		if (model.equals("��ͨ��λģ��")) {
			TableInfo[] infos = new TableInfo[25];
			infos[0] = new TableInfo("��ҳ", "", false);
			infos[1] = new TableInfo("��Чҳ", "", false);
			infos[2] = new TableInfo("���ɱ�", "", false);
			infos[3] = new TableInfo("����ͼ", "", false);
			infos[4] = new TableInfo("ʽ����", "", false);
			infos[5] = new TableInfo("Ϳ��", "1", true);
			infos[6] = new TableInfo("������", "1", true);
			infos[7] = new TableInfo("��ĸ��", "1", true);
			infos[8] = new TableInfo("��˨��", "1", true);
			infos[9] = new TableInfo("�㺸-PSW", "1", true);
			infos[10] = new TableInfo("�㺸-RSW", "1", true);
			infos[11] = new TableInfo("�㺸-MSW", "1", true);
			infos[12] = new TableInfo("�㺸-SSW", "1", true);
			infos[13] = new TableInfo(" PSW", "", false);
			infos[14] = new TableInfo("RSW����", "", false);
			infos[15] = new TableInfo("RSW�ŷ�", "", false);
			infos[16] = new TableInfo("��ʱ����", "1", true);
			infos[17] = new TableInfo("������ҵ", "1", true);
			infos[18] = new TableInfo("HEMָʾ", "1", true);
			infos[19] = new TableInfo("������װ", "1", true);
			infos[20] = new TableInfo("װ��", "1", true);
			infos[21] = new TableInfo("���ͳ�Ʊ�", "", false);
			infos[22] = new TableInfo("NEPID��װ", "1", true);
			infos[23] = new TableInfo("����", "1", true);
			infos[24] = new TableInfo("�����߼��", "1", true);
			tableViewer.setInput(infos);
		} else if(model.equals("VIN����ģ��")) {
			TableInfo[] infos = new TableInfo[5];
			infos[0] = new TableInfo("��ҳ", "", false);
			infos[1] = new TableInfo("��Чҳ", "", false);
			infos[2] = new TableInfo("���ɱ�", "", false);
			infos[3] = new TableInfo("����ͼ", "", false);
			infos[4] = new TableInfo("�����ҵ", "", false);
			tableViewer.setInput(infos);
		}else {
			TableInfo[] infos = new TableInfo[14];
			infos[0] = new TableInfo("��ҳ", "", false);
			infos[1] = new TableInfo("��Чҳ", "", false);
			infos[2] = new TableInfo("���ɱ�", "", false);
			infos[3] = new TableInfo("����ͼ", "", false);
			infos[4] = new TableInfo("ë�̴�ĥ", "", false);
			infos[5] = new TableInfo("�����߼��", "", false);
			infos[6] = new TableInfo("��í", "", false);
			infos[7] = new TableInfo("װ��", "", false);
			infos[8] = new TableInfo("��װ", "", false);
			infos[9] = new TableInfo("��ۼ��", "", false);
			infos[10] = new TableInfo("��Ӿ��", "", false);
			infos[11] = new TableInfo("��Ӿ�ξ�", "", false);
			infos[12] = new TableInfo("�������", "", false);
			infos[13] = new TableInfo("NEPID��װ", "", false);
			tableViewer.setInput(infos);
		}
	}

	// ��ѯ�����ѡ���ȡ�����Ϣ
	private ArrayList getSizeRule() {
		ArrayList rule = new ArrayList();
		try {
			AbstractAIFUIApplication app = AIFUtility.getCurrentApplication();
			TCSession session = (TCSession) app.getSession();
			TCPreferenceService preferenceService = session.getPreferenceService();

			String str = preferenceService.getPreferenceDescription("DFL9_get_version_information");
			if (str != null) {
				String[] values = preferenceService.getStringValues("DFL9_get_version_information");
				for (int i = 0; i < values.length; i++) {
					String value = values[i];
					rule.add(value);
				}
			}
			return rule;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return rule;
	}

	public static void main(String[] args) {
		SelectTemplate dialog1 = new SelectTemplate(new Shell(), SWT.SHELL_TRIM);
		dialog1.open();
	}
}
