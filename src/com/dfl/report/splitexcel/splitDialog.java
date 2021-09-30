package com.dfl.report.splitexcel;

import org.eclipse.swt.widgets.Dialog;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.wb.swt.SWTResourceManager;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Text;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;

public class splitDialog extends Dialog {

	protected Object result;
	protected Shell shell;
	private Text text;
	private Label label_1;
	public String altercode;

	/**
	 * Create the dialog.
	 * 
	 * @param parent
	 * @param style
	 */
	public splitDialog(Shell parent, int style) {
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
		shell.setImage(
				SWTResourceManager.getImage(splitDialog.class, "/com/dfl/report/imags/defaultapplication_16.png"));
		shell.setSize(450, 160);
		shell.setText("\u586B\u5199\u66F4\u6539\u5355\u53F7");

		Composite composite = new Composite(shell, SWT.NONE);
		composite.setBounds(0, 0, 444, 131);
		composite.setLayout(null);

		text = new Text(composite, SWT.BORDER);
		text.setBounds(77, 7, 136, 23);

		Label label = new Label(composite, SWT.NONE);
		label.setBounds(10, 10, 61, 17);
		label.setText("\u66F4\u6539\u5355\u53F7\uFF1A");

		Button button = new Button(composite, SWT.NONE);
		button.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				okKey();
			}
		});
		button.setBounds(266, 77, 80, 27);
		button.setText("\u786E\u5B9A");

		Button button_1 = new Button(composite, SWT.NONE);
		button_1.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				shell.dispose();
			}
		});
		button_1.setText("\u53D6\u6D88");
		button_1.setBounds(352, 77, 80, 27);

		label_1 = new Label(composite, SWT.NONE);
		label_1.setForeground(SWTResourceManager.getColor(SWT.COLOR_RED));
		label_1.setBounds(10, 44, 422, 27);
	}

	protected void okKey() {
		// TODO Auto-generated method stub
		if (text.getText().isEmpty()) {
			label_1.setText("ÇëÌîÐ´¸ü¸Äµ¥ºÅ£¡");
			return;
		}
		altercode = text.getText();
		shell.dispose();
	}
}
