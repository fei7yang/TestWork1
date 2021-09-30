package com.dfl.report.Fixturestylebook;

import org.eclipse.swt.widgets.Dialog;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.wb.swt.SWTResourceManager;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.SWT;
import org.eclipse.swt.layout.FillLayout;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.layout.FormLayout;
import org.eclipse.swt.layout.FormData;
import org.eclipse.swt.layout.FormAttachment;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.widgets.Text;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.events.VerifyListener;
import org.eclipse.swt.events.VerifyEvent;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;

public class UpdateConfirmDialog extends Dialog {

	protected Object result;
	protected Shell shell;
	//public String page1;
	public String updateflag = "1";//1：终止；2：不更新；3更新
	private Button btnCheckButton;
	private Button button;

	/**
	 * Create the dialog.
	 * @param parent
	 * @param style
	 */
	public UpdateConfirmDialog(Shell parent, int style) {
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
		shell.setImage(SWTResourceManager.getImage(UpdateConfirmDialog.class, "/com/dfl/report/dialog/defaultapplication_16.png"));
		shell.setSize(547, 295);
		shell.setText("\u662F\u5426\u66F4\u65B0\u6784\u6210\u90E8\u54C1\u4E00\u89C8\u9875\uFF1F");
		shell.setLayout(new FillLayout(SWT.HORIZONTAL));
		
		Composite composite = new Composite(shell, SWT.NONE);
		composite.setLayout(null);
		
		Button btnNewButton = new Button(composite, SWT.NONE);
		btnNewButton.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				Okkey();
			}
		});
		btnNewButton.setBounds(244, 210, 80, 27);
		btnNewButton.setText("\u786E\u5B9A");
		
		btnCheckButton = new Button(composite, SWT.CHECK);
		btnCheckButton.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				if(btnCheckButton.getSelection())
				{
					if(button.getSelection())
					{
						button.setSelection(false);
					}
				}
			}
		});
		btnCheckButton.setSelection(true);
		btnCheckButton.setBounds(32, 30, 437, 32);
		btnCheckButton.setText("\u4E0D\u66F4\u65B0   \u6CE8\uFF1A\u6784\u6210\u90E8\u54C1\u4E00\u89C8\u9875\u4EBA\u5DE5\u4FEE\u6539\u7684\u5185\u5BB9\u4E0D\u66F4\u65B0");
		
		button = new Button(composite, SWT.CHECK);
		button.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				if(button.getSelection())
				{
					if(btnCheckButton.getSelection())
					{
						btnCheckButton.setSelection(false);
					}
				}
			}
		});
		button.setText("\u66F4\u65B0       \u6CE8\uFF1A\u6784\u6210\u90E8\u54C1\u4E00\u89C8\u9875\u7CFB\u7EDF\u8F93\u51FA\u5185\u5BB9\u66F4\u65B0\u8986\u76D6");
		button.setBounds(32, 67, 437, 32);
		
		Label lblNewLabel = new Label(composite, SWT.WRAP);
		lblNewLabel.setBounds(120, 103, 352, 90);
		lblNewLabel.setText("\uFF08\u4F8B\u5982\uFF1A\u672C\u5DE5\u4F4D\u6709\u6765\u81EA\u522B\u7684\u4EA7\u7EBF\u7684\u603B\u6210\u6216\u4E0A\u4E00\u5DE5\u4F4D\u603B\u6210\u4E0A\u4EF6\uFF0C\u5982\u8FD9\u4E9B\u603B\u6210\u7CFB\u7EDF\u8F93\u51FA\u5230\u6784\u6210\u8868\u9875\uFF0C\u4EBA\u5DE5\u4FEE\u6B63\u8FC7\u5176\u4E0A\u4EF6\u987A\u5E8F\uFF0C\u672C\u6B21\u7CFB\u7EDF\u8FD8\u662F\u4F1A\u5C06\u5176\u5728\u6784\u6210\u8868\u9875Partlist\u90E8\u5206\u6700\u540E\u4E0A\u4EF6\u7684\u987A\u5E8F\u680F\u8F93\u51FA\uFF09");

	}

	protected void Okkey() {
		// TODO Auto-generated method stub				
		//page1 = text.getText();
	    if(btnCheckButton.getSelection())
	    {
	    	updateflag = "2";
	    }
	    if(button.getSelection())
	    {
	    	updateflag = "3";
	    }
		
		shell.dispose();
	}
	public static void main(String[] args) {
		UpdateConfirmDialog dialog1 = new UpdateConfirmDialog(new Shell(),SWT.SHELL_TRIM);
		dialog1.open();
		System.out.println(dialog1.updateflag);
	}
}
