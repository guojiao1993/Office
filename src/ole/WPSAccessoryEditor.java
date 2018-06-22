package ole;

import org.eclipse.swt.SWT;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.events.SelectionListener;
import org.eclipse.swt.graphics.Image;
import org.eclipse.swt.graphics.ImageData;
import org.eclipse.swt.layout.FillLayout;
import org.eclipse.swt.widgets.*;
import org.eclipse.swt.widgets.FileDialog;
import org.eclipse.swt.widgets.Menu;
import org.eclipse.swt.widgets.MenuItem;

import java.awt.Button;
import java.awt.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map.Entry;
import java.util.Set;

/**
 * WPS附件管理类
 *
 * @author Administrator
 *
 */
public class WPSAccessoryEditor implements Runnable {

	public static final int HEIGHT = 800;
	public static final int WIDTH = 900;
	private WPSWrapper wpsWrapper;
	private Display display;
	private Shell shell;
	private Button saveButton;
	private String filePath;
	private boolean editable;
	private int showmenu;

	private String photoPath;
	private HashMap<String, String> headValueMap;
	private HashMap<Integer, ArrayList<String[]>> bodyValueMap;

	public WPSAccessoryEditor(String filePath, boolean editable, int showmenu) {
		this.filePath = filePath;
		this.editable = editable;
		this.showmenu = showmenu;
		new Thread(this).start();
	}

	public WPSAccessoryEditor(String photoPath, HashMap<String, String> headValueMap,
							  HashMap<Integer, ArrayList<String[]>> bodyValueMap,
							  String filePath, boolean editable, int showmenu) {
		this.photoPath = photoPath;
		this.headValueMap = headValueMap;
		this.bodyValueMap = bodyValueMap;

		this.filePath = filePath;
		this.editable = editable;
		this.showmenu = showmenu;
		new Thread(this).start();
	}

	public void setIcon() {
		ImageData id = new ImageData(
				WPSAccessoryEditor.class.getResourceAsStream("logo.png"));
		shell.setImage(new Image(display, id));
	}

	public void hileBorder() {
		wpsWrapper.hileBorder();
	}

	@Override
	public void run() {
		S ldt = new S("加载中，请稍后……");
		display = new Display();
		shell = new Shell(display);

		shell.setText("OLE");
//		setIcon();

		shell.setLayout(new FillLayout());
		shell.layout();
		shell.setSize(WIDTH, HEIGHT);

		wpsWrapper = new WPSWrapper(shell, SWT.ON_TOP | SWT.SYSTEM_MODAL
				| SWT.NO_TRIM, filePath);

		int x = (Toolkit.getDefaultToolkit().getScreenSize().width - WIDTH) / 2;
		int y = (Toolkit.getDefaultToolkit().getScreenSize().height - HEIGHT) / 2;
		shell.setLocation(x, y);

		// 写入表头照片
		if(photoPath != null) {
			wpsWrapper.setPicture("photo", photoPath);
		}
		// 写入表头内容
		if(headValueMap != null) {
			Set<Entry<String, String>> headValueEntrySet = headValueMap.entrySet();
			for (Entry<String, String> entry : headValueEntrySet) {
				wpsWrapper.setFiled(entry.getKey(), entry.getValue());
			}
		}

		// 写入表体内容
		if(bodyValueMap != null) {
			Set<Entry<Integer, ArrayList<String[]>>> bodyValueEntrySet = bodyValueMap.entrySet();
			for (Entry<Integer, ArrayList<String[]>> entry : bodyValueEntrySet) {
				Integer key = entry.getKey();
				ArrayList<String[]> valueList = entry.getValue();
				if(valueList.size() > 0) {
					wpsWrapper.addTableItems(key, false, valueList.get(0));
					for (int i = 1; i < valueList.size(); i++) {
						wpsWrapper.addTableItems(key, true, valueList.get(i));
					}
				}
			}
		}

		// 增加另存为的菜单项
		addSaveOtherMenu();

		// 对文档加锁
		wpsWrapper.lock();

		shell.open();
		ldt.dispose();
		while (!shell.isDisposed()) {
			if (!display.readAndDispatch()) {
				display.sleep();
			}
		}
	}

	public WPSWrapper getWpsWrapper() {
		return wpsWrapper;
	}

	public void setWpsWrapper(WPSWrapper wpsWrapper) {
		this.wpsWrapper = wpsWrapper;
	}

	public Display getDisplay() {
		return display;
	}

	public void setDisplay(Display display) {
		this.display = display;
	}

	public Shell getShell() {
		return shell;
	}

	public void setShell(Shell shell) {
		this.shell = shell;
	}

	public void addCustomizedButton() {
		saveButton = new Button();
		saveButton.setLabel("保存");
	}

	public void addMenu2print() {
		Menu menu = new Menu(shell, SWT.BAR);
		MenuItem file = new MenuItem(menu, SWT.CASCADE);
		file.setText("&打印");
		Menu filemenu = new Menu(shell, SWT.DROP_DOWN);
		file.setMenu(filemenu);
		final MenuItem printItem = new MenuItem(filemenu, SWT.CASCADE);
		printItem.setText("&打印");

		shell.setMenuBar(menu);

		printItem.addSelectionListener(new SelectionListener() {
			@Override
			public void widgetDefaultSelected(SelectionEvent arg0) {
			}

			@Override
			public void widgetSelected(SelectionEvent arg0) {
				// 显示报文提示框
				wpsWrapper.printDialog();
			}
		});
	}

	public void addMenu() {
		Menu menu = new Menu(shell, SWT.BAR);
		MenuItem file = new MenuItem(menu, SWT.CASCADE);
		file.setText("&文件");
		Menu filemenu = new Menu(shell, SWT.DROP_DOWN);
		file.setMenu(filemenu);
		final MenuItem saveItem = new MenuItem(filemenu, SWT.CASCADE);
		saveItem.setText("&保存");
		shell.setMenuBar(menu);

		if (!editable)
			saveItem.setEnabled(false);
		else
			saveItem.setEnabled(true);
		saveItem.addSelectionListener(new SelectionListener() {
			@Override
			public void widgetDefaultSelected(SelectionEvent arg0) {

			}

			@Override
			public void widgetSelected(SelectionEvent arg0) {
				// 显示报文提示框
				InformationDialog dialog = new InformationDialog("文件保存中，请稍后");
				wpsWrapper.saveDoc();
				dialog.dispose();
			}
		});
	}


	public void addSaveOtherMenu() {
		Menu menu = new Menu(shell, SWT.BAR);
		MenuItem file = new MenuItem(menu, SWT.CASCADE);
		file.setText("&文件");

		Menu filemenu = new Menu(shell, SWT.DROP_DOWN);
		final MenuItem saveItem = new MenuItem(filemenu, SWT.CASCADE);
		saveItem.setText("&另存为");
		final MenuItem printItem = new MenuItem(filemenu, SWT.CASCADE);
		printItem.setText("&打印");
		final MenuItem quitItem = new MenuItem(filemenu, SWT.CASCADE);
		quitItem.setText("&退出");

		file.setMenu(filemenu);
		shell.setMenuBar(menu);

		saveItem.addSelectionListener(new SelectionListener() {
			@Override
			public void widgetDefaultSelected(SelectionEvent arg0) {

			}
			@Override
			public void widgetSelected(SelectionEvent arg0) {
				FileDialog filedlg = new FileDialog(shell, SWT.OPEN);
				// 设置文件对话框的标题
				filedlg.setText("文件选择框");
				// 设置过滤的后缀类型和名称
				filedlg.setFilterExtensions(new String[] { "*.doc" });
				filedlg.setFilterNames(new String[] { "Microsoft Word 97-2003 文档(*.doc)" });
				// 打开文件对话框，返回选中文件的绝对路径
				String otherFilePath = filedlg.open();
				if(otherFilePath != null) {
					if(!otherFilePath.endsWith(".doc")) {
						otherFilePath += ".pdf";
					}
					// 另存为新文档
					wpsWrapper.saveDoc(otherFilePath);
					// 完成后进行提示，一秒后自动消失
					InformationDialog dialog = new InformationDialog("文件保存成功");
					try {
						Thread.sleep(1000);
					} catch (InterruptedException e) {
						e.printStackTrace();
					}
					dialog.dispose();
				}
			}
		});

		quitItem.addSelectionListener(new SelectionListener() {
			@Override
			public void widgetDefaultSelected(SelectionEvent arg0) {

			}
			@Override
			public void widgetSelected(SelectionEvent arg0) {
				shell.dispose();
			}
		});

		printItem.addSelectionListener(new SelectionListener() {
			@Override
			public void widgetDefaultSelected(SelectionEvent arg0) {
			}
			@Override
			public void widgetSelected(SelectionEvent arg0) {
				wpsWrapper.printDialog();
			}
		});
	}

}
