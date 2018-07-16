
package ole2;

import java.io.File;
import java.util.HashMap;

import org.eclipse.swt.SWT;
import org.eclipse.swt.layout.FillLayout;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.OleClientSite;
import org.eclipse.swt.ole.win32.OleFrame;
import org.eclipse.swt.ole.win32.Variant;
import org.eclipse.swt.widgets.Composite;

/**
 * WPS������װ��
 * 
 * @author Administrator
 * 
 */
public class WPSWrapper extends Composite {
	private int width;
	private int height;
	private int islock;// WPS�ļ��Ƿ����
	private boolean showeditor;// �Ƿ���ʾ�༭��

	public WPSWrapper(Composite parent, int width, int height, String filePath,
			int islock, boolean showeditor) {
		super(parent, SWT.ON_TOP | SWT.SYSTEM_MODAL | SWT.NO_TRIM);
		this.filePath = filePath;
		this.width = width;
		this.height = height;
		this.islock = islock;
		this.showeditor = showeditor;
		initGUI();
	}

	private OleFrame oleFrame;
	private OleAutomation doc;
	private OleClientSite site;
	private String filePath;
	private File file;

	private void initGUI() {
		try {
			FillLayout thisLayout = new FillLayout(
					SWT.HORIZONTAL);
			this.setLayout(thisLayout);
			this.setSize(width, height);
			{
				oleFrame = new OleFrame(this, SWT.NONE);
				{
					file = new File(filePath);
					site = new OleClientSite(
							oleFrame, SWT.NONE,
							"kwps.document", file);
					site.setBounds(0, 0, width, height);
					site.doVerb(org.eclipse.swt.ole.win32.OLE.OLEIVERB_UIACTIVATE);
					//�Ƿ����ر༭����
					if (showeditor) {
						hileBorder();
					}
					// 0����������1�����ı�����Ա༭��2��ֻ��
					switch (islock) {
					case 0:
						break;
					case 1:
						lock();
						break;
					case 2:
						lockentirely();
						break;
					}
				}
			}
	        setFiled("t1", "��������");
			this.layout();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	/**
	 * ֻ��ģʽ�������ı������ݼ�ֵ��
	 * 
	 * @param name
	 * @param value
	 */
	public void setFiledwithreadonly(String name, String value){
		unlock();
		setFiled(name, value);
		lockentirely();
	}
	/**
	 * �����ı������ݼ�ֵ��
	 * 
	 * @param name
	 * @param value
	 */
	public void setFiled(String name, String value) {
		String[] split = value.split( "\n");
		String back="";
		if(split!=null&&split.length>1){
			for(String a :split){
				String ar=a+"\r";
				back+=ar;
			}
		}else{
			back=value;
		}
		try {
			doc = new OleAutomation(site);

			int FormFields[] = doc.getIDsOfNames(new String[] { "FormFields" });
			Variant bmk = doc.getProperty(FormFields[0]);
			OleAutomation bmkBook = bmk.getAutomation();

			int count[] = bmkBook.getIDsOfNames(new String[] { "Count" });
			Variant rgCnt = bmkBook.getProperty(count[0]);
			int fieldCount = rgCnt.getInt();

			for (int i = 1; i <= fieldCount; i++) {
				int item[] = bmkBook.getIDsOfNames(new String[] { "Item" });
				Variant it = bmkBook.invoke(item[0],
						new Variant[] { new Variant(i) });
				OleAutomation itAu = it.getAutomation();
				int text1[] = itAu.getIDsOfNames(new String[] { "Result" });

				int rng1[] = itAu.getIDsOfNames(new String[] { "Name" });
				Variant rgVa1 = itAu.getProperty(rng1[0]);

				String sText = rgVa1.getString();

				System.out.println(sText);
				System.out.println(name);
				if (sText.equals(name)) {
					int rng2[] = itAu.getIDsOfNames(new String[] { "Select" });
					itAu.invoke(rng2[0]);
					itAu.setProperty(text1[0], new Variant[] { new Variant(
							back) });
				}
			}
			doc.dispose();
		} catch (org.eclipse.swt.SWTException e) {
			String str = "OleClientSite:" + e.toString();
			System.out.println(str);
			return;
		}
	}

	/**
	 * ���ͼƬˮӡ
	 * 
	 * @param path
	 * @param x
	 * @param y
	 * @param width
	 * @param height
	 */
	public void addPicWaterSet(String path, int x, int y, int width, int height) {
		doc = new OleAutomation(site);

		int FormFields[] = doc.getIDsOfNames(new String[] { "Sections" });
		Variant bmk = doc.getProperty(FormFields[0]);
		OleAutomation bmkBook = bmk.getAutomation();

		int item[] = bmkBook.getIDsOfNames(new String[] { "Item" });
		Variant it = bmkBook.invoke(item[0], new Variant[] { new Variant(1) });
		OleAutomation itAu = it.getAutomation();

		int rng[] = itAu.getIDsOfNames(new String[] { "Headers" });
		Variant rgVa = itAu.getProperty(rng[0]);
		OleAutomation rgAu = rgVa.getAutomation();
		int item1[] = rgAu.getIDsOfNames(new String[] { "Item" });
		Variant it1 = rgAu.invoke(item1[0], new Variant[] { new Variant(1) });
		OleAutomation itAu1 = it1.getAutomation(); // ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary)

		int rng3[] = itAu1.getIDsOfNames(new String[] { "Range" });
		rgVa = itAu1.getProperty(rng3[0]);
		OleAutomation rg = rgVa.getAutomation();

		int text1[] = itAu1.getIDsOfNames(new String[] { "Shapes" });
		rgVa = itAu1.getProperty(text1[0]);
		OleAutomation rgAu2 = rgVa.getAutomation();

		int org[] = rgAu2.getIDsOfNames(new String[] { "AddPicture" });

		Variant[] rgvarg = new Variant[8];
		rgvarg[0] = new Variant(path);
		rgvarg[1] = new Variant(false);
		rgvarg[2] = new Variant(true);
		rgvarg[3] = new Variant(x);
		rgvarg[4] = new Variant(y);
		rgvarg[5] = new Variant(width);
		rgvarg[6] = new Variant(height);
		rgvarg[7] = new Variant(rg);
		rgAu2.invoke(org[0], rgvarg);
	}

	/**
	 * ���ر߿�
	 */
	public void hileBorder() {
		doc = new OleAutomation(site);

		int selection[] = doc.getIDsOfNames(new String[] { "Parent" });
		Variant se = doc.getProperty(selection[0]);
		OleAutomation seOle = se.getAutomation();

		int addin[] = seOle.getIDsOfNames(new String[] { "AddIns" });
		Variant se2 = seOle.getProperty(addin[0]);
		OleAutomation seOle3 = se2.getAutomation();

		String url = WPSWrapper.class.getClassLoader()
				.getResource("template.docx").getPath();
		System.out.println(url);
		int addin1[] = seOle3.getIDsOfNames(new String[] { "Add" });
		Variant it = seOle3.invoke(addin1[0],
				new Variant[] { new Variant(url) });
	}

	// ֱ�Ӵ�ӡ
	public void print() {
		doc = new OleAutomation(site);
		int FormFields[] = doc.getIDsOfNames(new String[] { "PrintOut" });
		Variant it = doc.invoke(FormFields[0]);
	}

	/**
	 * ��ӱ������-�Ƚ���--����ӱ��--������
	 * 
	 * @param index
	 *            �������
	 * @param append
	 *            ׷�ӻ�������ģʽ
	 * @param items
	 *            ����
	 */
	public void addTableItemswithlock(int index, boolean append, String[] items) {
		unlock();
		addTableItems(index, append, items);
		lock();
	}

	/**
	 * ��ӱ������
	 * 
	 * @param index
	 *            �������
	 * @param append
	 *            ׷�ӻ�������ģʽ
	 * @param items
	 *            ����
	 */
	public void addTableItems(int index, boolean append, String[] items) {
		try {
			doc = new OleAutomation(site);

			int tables[] = doc.getIDsOfNames(new String[] { "Tables" });
			Variant tbl = doc.getProperty(tables[0]);
			OleAutomation tablesOle = tbl.getAutomation();

			int count[] = tablesOle.getIDsOfNames(new String[] { "Count" });
			Variant rgCnt = tablesOle.getProperty(count[0]);
			int tableCnt = rgCnt.getInt();

			if (tableCnt < index)
				return;

			int tableItmes[] = tablesOle.getIDsOfNames(new String[] { "Item" });
			Variant tableItmeVar = tablesOle.invoke(tableItmes[0],
					new Variant[] { new Variant(index) });
			OleAutomation currentTable = tableItmeVar.getAutomation();

			int rows[] = currentTable.getIDsOfNames(new String[] { "Rows" });
			Variant rowVar = currentTable.getProperty(rows[0]);
			OleAutomation rowOle = rowVar.getAutomation();
			// �����׷��ģʽ������ɾ���ɵ����ݣ����˱�����
			if (!append) {

				int[] rowCnt = rowOle.getIDsOfNames(new String[] { "Count" });
				Variant rowCntVar = rowOle.getProperty(rowCnt[0]);
				int rowsCnt = rowCntVar.getInt();
				for (int i = 1; i < rowsCnt; i++) {
					int item[] = rowOle.getIDsOfNames(new String[] { "Item" });
					Variant it = rowOle.invoke(item[0],
							new Variant[] { new Variant(2) });
					currentTable = it.getAutomation();

					int deleteFunc[] = currentTable
							.getIDsOfNames(new String[] { "Delete" });
					currentTable.invoke(deleteFunc[0], new Variant[] {});
				}
			}

			int addFunc[] = rowOle.getIDsOfNames(new String[] { "Add" });
			Variant currentRowVar = rowOle.invoke(addFunc[0], new Variant[] {});

			OleAutomation currentRowOle = currentRowVar.getAutomation();
			int select[] = currentRowOle
					.getIDsOfNames(new String[] { "Select" });
			currentRowOle.invoke(select[0]);

			for (int i = 0; items != null && i < items.length; i++) {

				int cells[] = currentRowOle
						.getIDsOfNames(new String[] { "Cells" });
				Variant crVar = currentRowOle.getProperty(cells[0]);
				;
				OleAutomation crOle = crVar.getAutomation();

				int cellItmes[] = crOle.getIDsOfNames(new String[] { "Item" });
				Variant cellItmeVar = crOle.invoke(cellItmes[0],
						new Variant[] { new Variant(i + 1) });
				OleAutomation cellItmeVarOle = cellItmeVar.getAutomation();

				int ranges[] = cellItmeVarOle
						.getIDsOfNames(new String[] { "Range" });
				Variant rangesVar = cellItmeVarOle.getProperty(ranges[0]);
				OleAutomation rangesOle = rangesVar.getAutomation();
				int text1[] = rangesOle.getIDsOfNames(new String[] { "Text" });
				rangesOle.setProperty(text1[0], new Variant[] { new Variant(
						items[i]) });
			}
			doc.dispose();

		} catch (org.eclipse.swt.SWTException e) {
			String str = "OleClientSite:" + e.toString();
			System.out.println(str);
			return;
		}

	}

	/**
	 * ȡ������
	 */
	public void unlock() {
		doc = new OleAutomation(site);
		int protect[] = doc.getIDsOfNames(new String[] { "unprotect" });
		doc.invoke(protect[0],new Variant[] {new Variant("B26") });

	}

	/**
	 * ���ı�����Ա༭
	 */
	public void lock() {
		doc = new OleAutomation(site);
		int[] protect = doc.getIDsOfNames(new String[] { "Protect" });
		doc.invoke(protect[0], new Variant[] { new Variant(2),
				new Variant(true),new Variant("B26") });
	}

	/**
	 * WPSֻ��
	 */
	public void lockentirely() {
		doc = new OleAutomation(site);
		int protect[] = doc.getIDsOfNames(new String[] { "protect" });
		doc.invoke(protect[0], new Variant[] { new Variant(1),
				new Variant(true),new Variant("B26") });

	}

	/**
	 * �����ļ�
	 */
	public void saveDoc() {
		// ������ļ����޸Ĺ�
		if (site.isDirty()) {
			// ����һ����ʱ�ļ�
			File tempFile = new File(filePath + ".tmp");
			// site
			file.renameTo(tempFile);
			if (site.save(file, true))
				tempFile.delete();
			else
				tempFile.renameTo(file);
		}
	}

	/**
	 * ��ӡ�ļ��Ի���
	 */
	public void printDialog() {
		doc = new OleAutomation(site);

		int selection[] = doc.getIDsOfNames(new String[] { "Parent" });
		Variant se = doc.getProperty(selection[0]);
		OleAutomation seOle = se.getAutomation();

		int dialogs[] = seOle.getIDsOfNames(new String[] { "Dialogs" });
		Variant bmk = seOle.getProperty(dialogs[0]);
		OleAutomation bmkBook = bmk.getAutomation();

		int item[] = bmkBook.getIDsOfNames(new String[] { "Item" });
		Variant it = bmkBook.invoke(item[0], new Variant[] { new Variant(88) });
		OleAutomation itAu = it.getAutomation();

		int item1[] = itAu.getIDsOfNames(new String[] { "Show" });
		it = itAu.invoke(item1[0]);

	}
}
