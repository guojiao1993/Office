package ole2;

import java.awt.Point;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import ole.WPSWrapper;
import org.eclipse.swt.SWT;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.events.SelectionListener;
import org.eclipse.swt.graphics.Image;
import org.eclipse.swt.graphics.ImageData;
import org.eclipse.swt.internal.win32.OS;
import org.eclipse.swt.layout.FillLayout;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Menu;
import org.eclipse.swt.widgets.MenuItem;
import org.eclipse.swt.widgets.Shell;

public class WPSDialog implements Runnable {

    public int width;

    public int height;

    private WPSWrapper wpsWrapper;

    private Display display;

    private Shell shell;

    private String filePath;

    private Point location;
    
    private int  islock;//0����������1�����ı�����Ա༭��2��ֻ��

    private boolean isOpen = false;

    public WPSDialog(String filePath, Point pt, int width, int height,int islock) {
        this.filePath = filePath;
        this.width = width;
        this.height = height;
        this.location = pt;
        this.islock=islock;
        new Thread(this).start();
    }

    public void setIcon() {
        ImageData id = new ImageData(WPSDialog.class.getResourceAsStream("logo.png"));
        this.shell.setImage(new Image(this.display, id));
    }

    public void hileBorder() {
        this.wpsWrapper.hileBorder();
    }

    @Override
    public void run() {
        this.display = new Display();
        this.shell = new Shell(this.display);
        this.shell.setText("BUCG");
        this.shell.setLayout(new FillLayout());
        this.shell.layout();
        this.shell.setSize(this.width, this.height);
        this.setIcon();
        this.wpsWrapper = new WPSWrapper(this.shell, this.width, this.height, this.filePath,islock,true);
        OS.SetWindowPos(this.shell.handle, OS.HWND_NOTOPMOST, this.location.x, this.location.y, this.width,
                this.height, SWT.NULL);
        this.shell.open();
        while (!this.shell.isDisposed()) {
            if (!this.display.readAndDispatch()) {
                this.display.sleep();
            }
        }

    }

    public WPSWrapper getWpsWrapper() {
        return this.wpsWrapper;
    }

    public void setWpsWrapper(WPSWrapper wpsWrapper) {
        this.wpsWrapper = wpsWrapper;
    }

    public Display getDisplay() {
        return this.display;
    }

    public void setDisplay(Display display) {
        this.display = display;
    }

    public Shell getShell() {
        return this.shell;
    }

    public void setShell(Shell shell) {
        this.shell = shell;
    }

    public void addMenu() {
        Menu menu = new Menu(this.shell, SWT.BAR);
        MenuItem file = new MenuItem(menu, SWT.CASCADE);
        file.setText("&�ļ�");
        Menu filemenu = new Menu(this.shell, SWT.DROP_DOWN);
        file.setMenu(filemenu);
        final MenuItem saveItem = new MenuItem(filemenu, SWT.CASCADE);
        saveItem.setText("&����");
        this.shell.setMenuBar(menu);

        saveItem.addSelectionListener(new SelectionListener() {
            @Override
            public void widgetDefaultSelected(SelectionEvent arg0) {
            }

            @Override
            public void widgetSelected(SelectionEvent arg0) {
                WPSDialog.this.wpsWrapper.saveDoc();
            }
        });
    }

    public void setLocation(final Point pt) {
        if (this.display != null && this.shell != null) {
            this.display.asyncExec(new Runnable() {
                @Override
                public void run() {
                    OS.SetWindowPos(WPSDialog.this.shell.handle, OS.HWND_NOTOPMOST, pt.x, pt.y, WPSDialog.this.width,
                            WPSDialog.this.height, SWT.NULL);
                }
            });
        }
    }

    public boolean isOpen() {
        return this.isOpen;
    }

    public void setVisable(final boolean flag) {

        if (this.display != null && this.shell != null) {
            this.display.asyncExec(new Runnable() {
                @Override
                public void run() {
                    WPSDialog.this.shell.setVisible(flag);
                    WPSDialog.this.isOpen = flag;
                }
            });
        }
    }

    public void destroy() {
        if (this.display != null && this.shell != null) {
            this.display.asyncExec(new Runnable() {
                @Override
                public void run() {
                    WPSDialog.this.shell.dispose();
                }
            });
        }
    }

    public void saveDoc() {
        if (this.display != null && this.shell != null) {
            this.display.asyncExec(new Runnable() {
                @Override
                public void run() {
                    if (WPSDialog.this.wpsWrapper != null) {
                        WPSDialog.this.wpsWrapper.saveDoc();
                    }
                }
            });
        }

    }

    public void printDialog() {
        if (this.display != null && this.shell != null) {
            this.display.asyncExec(new Runnable() {
                @Override
                public void run() {
                    if (WPSDialog.this.wpsWrapper != null) {
                        WPSDialog.this.wpsWrapper.printDialog();
                    }
                }
            });
        }

    }

    public void setFields(final Map<String, String> pairs) {
        if (this.display != null && this.shell != null) {
            this.display.asyncExec(new Runnable() {
                @Override
                public void run() {
                    if (WPSDialog.this.wpsWrapper != null) {
                        try {
                            Iterator<Map.Entry<String, String>> entries = pairs.entrySet().iterator();
                            while (entries.hasNext()) {
                                Map.Entry<String, String> entry = entries.next();
                                WPSDialog.this.wpsWrapper.setFiled(entry.getKey(), entry.getValue());
                            }
                        }
                        catch (Exception e) {
//                            WPSDialog.this.ncDialog.showErrorMsg("������ʾ", "�����ı���ʧ��:" + e.getMessage());
                            e.printStackTrace();
                        }
                    }
                }
            });
        }
    }
    public void setFiledwithreadonly(final Map<String, String> pairs) {
        if (this.display != null && this.shell != null) {
            this.display.asyncExec(new Runnable() {
                @Override
                public void run() {
                    if (WPSDialog.this.wpsWrapper != null) {
                        try {
                            Iterator<Map.Entry<String, String>> entries = pairs.entrySet().iterator();
                            while (entries.hasNext()) {
                                Map.Entry<String, String> entry = entries.next();
                                WPSDialog.this.wpsWrapper.setFiledwithreadonly(entry.getKey(), entry.getValue());
                            }
                        }
                        catch (Exception e) {
//                            WPSDialog.this.ncDialog.showErrorMsg("������ʾ", "�����ı���ʧ��:" + e.getMessage());
                            e.printStackTrace();
                        }
                    }
                }
            });
        }
    }
    public void addPicWaterSet(final String path, final int x, final int y, final int width, final int height) {
        if (this.display != null && this.shell != null) {
            this.display.asyncExec(new Runnable() {
                @Override
                public void run() {
                    if (WPSDialog.this.wpsWrapper != null) {
                        WPSDialog.this.wpsWrapper.addPicWaterSet(path, x, y, width, height);
                    }
                }
            });
        }
    }

    public void addTableItems(final int index, final boolean append, final String[] items) {
        if (this.display != null && this.shell != null) {
            this.display.asyncExec(new Runnable() {
                @Override
                public void run() {
                    if (WPSDialog.this.wpsWrapper != null) {
                        WPSDialog.this.wpsWrapper.addTableItems(index, append, items);
                    }
                }
            });
        }
    }
    public void addTableItemswithlock(final int index, final boolean append, final String[] items) {
        if (this.display != null && this.shell != null) {
            this.display.asyncExec(new Runnable() {
                @Override
                public void run() {
                    if (WPSDialog.this.wpsWrapper != null) {
                        WPSDialog.this.wpsWrapper.addTableItemswithlock(index, append, items);
                    }
                }
            });
        }
    }
    public void lock() {
        if (this.display != null && this.shell != null) {
            this.display.asyncExec(new Runnable() {
                @Override
                public void run() {
                    if (WPSDialog.this.wpsWrapper != null) {
                        WPSDialog.this.wpsWrapper.lock();
                    }
                }
            });
        }
    }
    public void unlock() {
        if (this.display != null && this.shell != null) {
            this.display.asyncExec(new Runnable() {
                @Override
                public void run() {
                    if (WPSDialog.this.wpsWrapper != null) {
                        WPSDialog.this.wpsWrapper.unlock();
                    }
                }
            });
        }
    }

    public void testActionCallback() {
        System.out.println("1111111111111111111111111");
    }
}
