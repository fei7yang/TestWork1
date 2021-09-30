package com.dfl.report.workschedule;

import org.eclipse.swt.widgets.Dialog;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.Composite;

import java.util.ArrayList;

import org.eclipse.swt.SWT;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.widgets.Label;
import swing2swt.layout.BorderLayout;
import org.eclipse.swt.widgets.Combo;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.widgets.Button;
import org.eclipse.wb.swt.SWTResourceManager;

import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCSession;

import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.events.ModifyListener;
import org.eclipse.swt.events.ModifyEvent;

public class EditionDialog extends Dialog {

	protected Object result;
	protected Shell shell;
	private Combo combo;
	public String Edition;
	private Label label_1;
	private Combo combo_1;
	private String lastedition;
    private boolean flag = false;
    private ArrayList rule;
	/**
	 * Create the dialog.
	 * 
	 * @param parent
	 * @param style
	 * @param lastEdition 
	 * @param rule 
	 */
	public EditionDialog(Shell parent, int style, String lastEdition, ArrayList rule) {
		super(parent, style);
		setText("SWT Dialog");
		lastedition = lastEdition;
		this.rule = rule;
	}
	public EditionDialog(Shell parent, int style, boolean flag, ArrayList rule) {
		super(parent, style);
		setText("SWT Dialog");
		this.flag = flag;
		this.rule = rule;
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
		shell.setImage(SWTResourceManager.getImage(EditionDialog.class,
				"/com/dfl/report/imags/defaultapplication_16.png"));
		shell.setSize(532, 176);
		shell.setText("\u9009\u62E9\u7248\u6B21");
		shell.setLayout(new BorderLayout(0, 0));

		Composite composite = new Composite(shell, SWT.NONE);
		composite.setLayoutData(BorderLayout.NORTH);
		composite.setLayout(null);

		Label label_2 = new Label(composite, SWT.NONE);
		label_2.setBounds(5, 9, 60, 17);
		label_2.setText("\u8F93\u51FA\u9636\u6BB5\uFF1A");

		combo_1 = new Combo(composite, SWT.NONE);
		combo_1.setBounds(70, 5, 108, 25);
		combo_1.addModifyListener(new ModifyListener() {
			public void modifyText(ModifyEvent arg0) {
				ComboValueChangeEvent();
			}
		});
		combo_1.add("SOP前");
		combo_1.add("SOP后");
		combo_1.select(0);

		Label label = new Label(composite, SWT.NONE);
		label.setBounds(199, 9, 36, 17);
		label.setText("\u7248\u6B21\uFF1A");

		combo = new Combo(composite, SWT.NONE);
		combo.setBounds(248, 6, 108, 25);

		Composite composite_1 = new Composite(shell, SWT.NONE);
		composite_1.setLayoutData(BorderLayout.SOUTH);
		composite_1.setLayout(null);

		// 确定事件
		Button button = new Button(composite_1, SWT.NONE);
		button.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				Okevent();
			}
		});
		button.setBounds(390, 5, 55, 27);
		button.setText("\u786E\u5B9A");

		// 取消事件
		Button button_1 = new Button(composite_1, SWT.NONE);
		button_1.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				shell.dispose();
			}
		});
		button_1.setBounds(451, 5, 55, 27);
		button_1.setText("\u53D6\u6D88");

		Composite composite_2 = new Composite(shell, SWT.NONE);
		composite_2.setLayoutData(BorderLayout.CENTER);

		label_1 = new Label(composite_2, SWT.NONE);
		label_1.setForeground(SWTResourceManager.getColor(SWT.COLOR_RED));
		label_1.setBounds(50, 10, 466, 25);

		Label label_3 = new Label(composite_2, SWT.NONE);
		label_3.setForeground(SWTResourceManager.getColor(SWT.COLOR_RED));
		label_3.setText("\u4E0A\u4E00\u6B21\u7684\u7248\u6B21\u662F\uFF1A" + lastedition);
		label_3.setBounds(45, 41, 471, 29);
		
		if(flag) {
			label_3.setVisible(false);
		}
	
		if (rule != null && rule.size() > 0) {
			for (int i = 0; i < rule.size(); i++) {
				combo.add((String) rule.get(i));
			}
		} else {
			label_1.setText("版次信息首选项(DFL9_get_version_information)未设置！");
		}

	}

	protected void ComboValueChangeEvent() {
		// TODO Auto-generated method stub
		if(combo_1.getText().equals("SOP前")) {
			if(combo!=null) {
				combo.removeAll();
				ArrayList list = getSizeRule();
				if (list != null && list.size() > 0) {
					for (int i = 0; i < list.size(); i++) {
						combo.add((String) list.get(i));
					}
				} else {
					label_1.setText("版次信息首选项(DFL9_get_version_information)未设置！");
				}
			}
			
		}
		if(combo_1.getText().equals("SOP后")) {
			if(combo!=null) {
				combo.removeAll();
			}		
		}
	}

	protected void Okevent() {
		// TODO Auto-generated method stub
		if (combo.getText().isEmpty()) {
			label_1.setText("请选择版次或填写版次！");
			return;
		}
		Edition = combo.getText();
		System.out.println("Edition:" + Edition);
		shell.dispose();
	}

	// 查询版次首选项，获取版次信息
	private ArrayList getSizeRule() {
		ArrayList rule = new ArrayList();
		try {
			AbstractAIFUIApplication app = AIFUtility.getCurrentApplication();
			TCSession session = (TCSession) app.getSession();
			TCPreferenceService preferenceService = session.getPreferenceService();

			String str = preferenceService.getPreferenceDescription("DFL9_get_version_information");
			if (str != null) {
				String[] values = preferenceService.getStringValues("DFL9_get_version_information");
				if(values!=null) {
					for (int i = 0; i < values.length; i++) {
						String value = values[i];
						rule.add(value);
					}
				}			
			}
			return rule;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return rule;
	}

	public static void main(String[] args) {
		EditionDialog dialog1 = new EditionDialog(new Shell(), SWT.SHELL_TRIM,"",null);
		dialog1.open();
	}
}
