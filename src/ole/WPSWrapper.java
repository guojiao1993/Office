package ole;

import org.eclipse.swt.SWT;
import org.eclipse.swt.layout.FillLayout;
import org.eclipse.swt.ole.win32.*;
import org.eclipse.swt.widgets.Composite;

import java.io.File;

/**
 * WPS包装类
 *
 * @author Administrator
 */
public class WPSWrapper extends Composite {
    private OleFrame oleFrame;
    private OleAutomation doc;
    private OleClientSite site;
    private String filePath;

    public WPSWrapper(Composite parent, int style, String filePath) {
        super(parent, style);
        this.filePath = filePath;
        initGUI();
    }

    public void onButtonClicked() {
        setFiled("test1", "abcdef");
        addTableItems(1, false, new String[]{"1", "2", "3", "4", "5"});
        addTableItems(2, true, new String[]{"1", "2", "3", "4", "5"});
    }

    private void initGUI() {
        try {
            FillLayout thisLayout = new FillLayout(org.eclipse.swt.SWT.HORIZONTAL);
            this.setLayout(thisLayout);
            this.setSize(229, 54);

            oleFrame = new OleFrame(this, SWT.NONE);

            site = new OleClientSite(oleFrame, org.eclipse.swt.SWT.NONE, "kwps.document", new File(filePath));
            site.setBounds(0, 0, 104, 54);
            site.doVerb(OLE.OLEIVERB_UIACTIVATE);

            this.layout();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 设置文本域的值
     *
     * @param name
     * @param value
     */
    public void setFiled(String name, String value) {
        try {
            doc = new OleAutomation(site);

            int FormFields[] = doc.getIDsOfNames(new String[]{"FormFields"});
            Variant bmk = doc.getProperty(FormFields[0]);
            OleAutomation bmkBook = bmk.getAutomation();

            int count[] = bmkBook.getIDsOfNames(new String[]{"Count"});
            Variant rgCnt = bmkBook.getProperty(count[0]);
            int fieldCount = rgCnt.getInt();

            for (int i = 1; i <= fieldCount; i++) {
                int item[] = bmkBook.getIDsOfNames(new String[]{"Item"});
                Variant it = bmkBook.invoke(item[0], new Variant[]{new Variant(i)});
                OleAutomation itAu = it.getAutomation();
                int text1[] = itAu.getIDsOfNames(new String[]{"Result"});

                int rng1[] = itAu.getIDsOfNames(new String[]{"Name"});
                Variant rgVa1 = itAu.getProperty(rng1[0]);
                String sText = rgVa1.getString();

                System.out.println(name + " -> " + sText);

                if (sText.equals(name)) {
                    int rng2[] = itAu.getIDsOfNames(new String[]{"Select"});
                    itAu.invoke(rng2[0]);
                    itAu.setProperty(text1[0], new Variant[]{new Variant(value)});
                    return;
                }
            }
            doc.dispose();
        } catch (org.eclipse.swt.SWTException e) {
            System.out.println("OleClientSite:" + e.toString());
        }
    }

    /**
     * 设置文本域的图片
     *
     * @param name
     * @param path
     */
    public void setPicture(String name, String path) {
        try {
            doc = new OleAutomation(site);
            // 得到应用
            Variant av = doc.getProperty(doc.getIDsOfNames(new String[]{"Application"})[0]);
            OleAutomation ao = av.getAutomation();
            // 得到选中区域
            Variant sv = ao.getProperty(ao.getIDsOfNames(new String[]{"Selection"})[0]);
            // 得到所有文本域
            Variant bmk = doc.getProperty(doc.getIDsOfNames(new String[]{"FormFields"})[0]);
            OleAutomation bmkBook = bmk.getAutomation();
            // 得到文本域的数量
            Variant rgCnt = bmkBook.getProperty(bmkBook.getIDsOfNames(new String[]{"Count"})[0]);
            int fieldCount = rgCnt.getInt();
            // 遍历所有文本域
            for (int i = 1; i <= fieldCount; i++) {
                // 得到文本域
                Variant it = bmkBook.invoke(bmkBook.getIDsOfNames(new String[]{"Item"})[0], new Variant[]{new Variant(i)});
                OleAutomation itAu = it.getAutomation();
                // 得到文本域的名称
                Variant rgVa1 = itAu.getProperty(itAu.getIDsOfNames(new String[]{"Name"})[0]);
                if (rgVa1.getString().equals(name)) {
                    // 选中该文本域
                    itAu.invoke(itAu.getIDsOfNames(new String[]{"Select"})[0]);
                    // 设置选中的区域只限于该文本域
                    OleAutomation svo = sv.getAutomation();
                    Variant vstart = svo.getProperty(svo.getIDsOfNames(new String[]{"Start"})[0]);
                    svo.invoke(svo.getIDsOfNames(new String[]{"SetRange"})[0], new Variant[]{vstart, vstart});
                    // 得到选中的位置信息
                    int iInfo[] = svo.getIDsOfNames(new String[]{"Information"});
                    Variant vx = svo.getProperty(iInfo[0], new Variant[]{new Variant(5)});
                    Variant vy = svo.getProperty(iInfo[0], new Variant[]{new Variant(6)});
                    // 得到选中区域的内联图形
                    Variant vInline = svo.getProperty(svo.getIDsOfNames(new String[]{"InlineShapes"})[0]);
                    OleAutomation oInline = vInline.getAutomation();
                    // 在选中区域插入图片
                    Variant[] rgvarg = new Variant[3];
                    rgvarg[0] = new Variant(path);
                    rgvarg[1] = new Variant(false);
                    rgvarg[2] = new Variant(true);
                    Variant voInline = oInline.invoke(oInline.getIDsOfNames(new String[]{"AddPicture"})[0], rgvarg);
                    OleAutomation oInlineRs = voInline.getAutomation();
                    // 得到插入后的图形
                    Variant vcts = oInlineRs.invoke(oInlineRs.getIDsOfNames(new String[]{"ConvertToShape"})[0]);
                    OleAutomation octs = vcts.getAutomation();
                    // 设置插入后的图形的位置
                    Variant vwidth = octs.getProperty(octs.getIDsOfNames(new String[]{"Width"})[0]);
                    Variant vheight = octs.getProperty(octs.getIDsOfNames(new String[]{"Height"})[0]);
                    octs.setProperty(octs.getIDsOfNames(new String[]{"Left"})[0],
                            new Variant[]{new Variant(vx.getFloat() - vwidth.getFloat() / 2)});
                    octs.setProperty(octs.getIDsOfNames(new String[]{"Top"})[0],
                            new Variant[]{new Variant(vy.getFloat() - vheight.getFloat() / 2)});
                    return;
                }
            }
            doc.dispose();
        } catch (org.eclipse.swt.SWTException e) {
            System.out.println("OleClientSite:" + e.toString());
        }
    }

    /**
     * 盖章
     *
     * @param name
     * @param value
     * @author dnb
     */
    public void setSeal(String name, String value) {
        try {
            doc = new OleAutomation(site);

            int aa[] = doc.getIDsOfNames(new String[]{"Application"});
            Variant av = doc.getProperty(aa[0]);
            OleAutomation ao = av.getAutomation();
            int s[] = ao.getIDsOfNames(new String[]{"Selection"});
            System.out.println("==setSeal===:" + s.length);
            Variant sv = ao.getProperty(s[0]);
            OleAutomation svo = sv.getAutomation();
            int gtp[] = svo.getIDsOfNames(new String[]{"Range"});
            Variant vr = svo.getProperty(gtp[0]);
            OleAutomation ovr = vr.getAutomation();

            int iss[] = ovr.getIDsOfNames(new String[]{"Start"});
            int ise[] = ovr.getIDsOfNames(new String[]{"End"});

            int FormFields[] = doc.getIDsOfNames(new String[]{"FormFields"});
            System.out.println("==setfiled===:" + FormFields.length);
            Variant bmk = doc.getProperty(FormFields[0]);
            OleAutomation bmkBook = bmk.getAutomation();

            int count[] = bmkBook.getIDsOfNames(new String[]{"Count"});
            Variant rgCnt = bmkBook.getProperty(count[0]);
            int fieldCount = rgCnt.getInt();
            System.out.println("==setfield count===:" + fieldCount);
            boolean b = false, e = false;
            int item[] = bmkBook.getIDsOfNames(new String[]{"Item"});
            for (int i = 1; i <= fieldCount; i++) {
                Variant it = bmkBook.invoke(item[0], new Variant[]{new Variant(i)});
                OleAutomation itAu = it.getAutomation();
                int text1[] = itAu.getIDsOfNames(new String[]{"Result"});

                int rng1[] = itAu.getIDsOfNames(new String[]{"Name"});
                Variant rgVa1 = itAu.getProperty(rng1[0]);
                String sText = rgVa1.getString();
                if (sText.equals(name)) {
                    int rng2[] = itAu.getIDsOfNames(new String[]{"Select"});
                    itAu.invoke(rng2[0]);

                    int istart[] = svo.getIDsOfNames(new String[]{"Start"});
                    Variant vstart = svo.getProperty(istart[0]);

                    int isr[] = svo.getIDsOfNames(new String[]{"SetRange"});
                    svo.invoke(isr[0], new Variant[]{vstart, vstart});

                    int iInfo[] = svo.getIDsOfNames(new String[]{"Information"});
                    Variant vx = svo.getProperty(iInfo[0], new Variant[]{new Variant(5)});
                    Variant vy = svo.getProperty(iInfo[0], new Variant[]{new Variant(6)});

                    int iInline[] = svo.getIDsOfNames(new String[]{"InlineShapes"});
                    Variant vInline = svo.getProperty(iInline[0]);
                    OleAutomation oInline = vInline.getAutomation();
                    Variant[] rgvarg = new Variant[3];
                    rgvarg[0] = new Variant(value);
                    rgvarg[1] = new Variant("False");
                    rgvarg[2] = new Variant("True");
                    int iAddPic[] = oInline.getIDsOfNames(new String[]{"AddPicture"});
                    Variant voInline = oInline.invoke(iAddPic[0], rgvarg);
                    OleAutomation oInlineRs = voInline.getAutomation();

                    int icts[] = oInlineRs.getIDsOfNames(new String[]{"ConvertToShape"});
                    Variant vcts = oInlineRs.invoke(icts[0]);
                    OleAutomation octs = vcts.getAutomation();

                    int iwidth[] = octs.getIDsOfNames(new String[]{"Width"});
                    Variant vwidth = octs.getProperty(iwidth[0]);
                    int iheight[] = octs.getIDsOfNames(new String[]{"Height"});
                    Variant vheight = octs.getProperty(iheight[0]);

                    int ileft[] = octs.getIDsOfNames(new String[]{"Left"});
                    octs.setProperty(ileft[0], new Variant[]{new Variant(vx.getFloat() - vwidth.getFloat() / 2)});
                    int itop[] = octs.getIDsOfNames(new String[]{"Top"});
                    octs.setProperty(itop[0], new Variant[]{new Variant(vy.getFloat() - vheight.getFloat() / 2)});
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
     * 增加水印
     *
     * @param path
     * @param x
     * @param y
     * @param width
     * @param height
     */
    public void addPicWaterSet(String path, int x, int y, int width, int height) {
        doc = new OleAutomation(site);

        int FormFields[] = doc.getIDsOfNames(new String[]{"Sections"});
        Variant bmk = doc.getProperty(FormFields[0]);
        OleAutomation bmkBook = bmk.getAutomation();

        int item[] = bmkBook.getIDsOfNames(new String[]{"Item"});
        Variant it = bmkBook.invoke(item[0], new Variant[]{new Variant(1)});
        OleAutomation itAu = it.getAutomation();

        int rng[] = itAu.getIDsOfNames(new String[]{"Headers"});
        Variant rgVa = itAu.getProperty(rng[0]);
        OleAutomation rgAu = rgVa.getAutomation();

        int item1[] = rgAu.getIDsOfNames(new String[]{"Item"});
        Variant it1 = rgAu.invoke(item1[0], new Variant[]{new Variant(1)});
        OleAutomation itAu1 = it1.getAutomation();

        int rng3[] = itAu1.getIDsOfNames(new String[]{"Range"});
        rgVa = itAu1.getProperty(rng3[0]);
        OleAutomation rg = rgVa.getAutomation();

        int text1[] = itAu1.getIDsOfNames(new String[]{"Shapes"});
        rgVa = itAu1.getProperty(text1[0]);
        OleAutomation rgAu2 = rgVa.getAutomation();

        int org[] = rgAu2.getIDsOfNames(new String[]{"AddPicture"});

        Variant[] rgvarg = new Variant[8];
        rgvarg[0] = new Variant(path);
        rgvarg[1] = new Variant(false);
        rgvarg[2] = new Variant(true);
        rgvarg[3] = new Variant(x);
        rgvarg[4] = new Variant(y);
        rgvarg[5] = new Variant(width);
        rgvarg[6] = new Variant(height);
        rgvarg[7] = new Variant(rg);
        Variant it2 = rgAu2.invoke(org[0], rgvarg);

        int item3[] = rgAu2.getIDsOfNames(new String[]{"Item"});
        it = rgAu2.invoke(item3[0], new Variant[]{new Variant(1)});
        OleAutomation itAu4 = it.getAutomation();

        int item4[] = itAu4.getIDsOfNames(new String[]{"PictureFormat"});
        it = itAu4.getProperty(item4[0]);
        OleAutomation itAu5 = it.getAutomation();

        int item5[] = itAu5.getIDsOfNames(new String[]{"Brightness"});
        itAu5.setProperty(item5[0], new Variant[]{new Variant(0.5)});

        int item6[] = itAu5.getIDsOfNames(new String[]{"Contrast"});
        itAu5.setProperty(item6[0], new Variant[]{new Variant(0.15)});

    }

    public void hileBorder() {
        doc = new OleAutomation(site);
        int cmdBars[] = doc.getIDsOfNames(new String[]{"CommandBars"});
        Variant cmdBarsVar = doc.getProperty(cmdBars[0]);
        OleAutomation cmdBarsOle = cmdBarsVar.getAutomation();

        int count[] = cmdBarsOle.getIDsOfNames(new String[]{"Count"});
        Variant cntVar = cmdBarsOle.getProperty(count[0]);
        int cmdBarsCnt = cntVar.getInt();
        System.out.println("cmdBarsCnt=" + cmdBarsCnt);

        int item[] = cmdBarsOle.getIDsOfNames(new String[]{"Item"});

        for (int i = 1; i < cmdBarsCnt; i++) {
            Variant barItmeVar = cmdBarsOle.getProperty(item[0], new Variant[]{new Variant("1")});
            OleAutomation barItmeOle = barItmeVar.getAutomation();

            int name[] = barItmeOle.getIDsOfNames(new String[]{"Name"});
            Variant barItemVar = barItmeOle.getProperty(name[0]);
            System.out.println(barItemVar.getString());

            int visible[] = barItmeOle.getIDsOfNames(new String[]{"Visible"});
            barItmeOle.setProperty(visible[0], new Variant[]{new Variant(false)});
        }

    }

    public void print() {
        doc = new OleAutomation(site);
        int FormFields[] = doc.getIDsOfNames(new String[]{"PrintOut"});
        Variant it = doc.invoke(FormFields[0]);
    }

    /**
     * 添加表格内容-先解锁--》添加表格--》上锁
     *
     * @param index  表格索引
     * @param append 追加或者新增模式
     * @param items  内容
     */
    public void addTableItemswithlock(int index, boolean append, String[] items) {
        unlock();
        addTableItems(index, append, items);
        lock();
    }

    /**
     * 取消限制
     */
    public void unlock() {
        doc = new OleAutomation(site);
        int protect[] = doc.getIDsOfNames(new String[]{"unprotect"});
        doc.invoke(protect[0]);
    }

    public void addTableItems(int index, boolean append, String[] items) {
        try {
            doc = new OleAutomation(site);
            // 得到所有的表格对象
            int tables[] = doc.getIDsOfNames(new String[]{"Tables"});
            Variant tbl = doc.getProperty(tables[0]);
            OleAutomation tablesOle = tbl.getAutomation();
            // 得到表格数量
            int count[] = tablesOle.getIDsOfNames(new String[]{"Count"});
            Variant rgCnt = tablesOle.getProperty(count[0]);
            int tableCnt = rgCnt.getInt();
            // 目标表格不存在，则返回
            if (tableCnt < index) {
                return;
            }
            // 得到要操作的表格对象
            int tableItmes[] = tablesOle.getIDsOfNames(new String[]{"Item"});
            Variant tableItmeVar = tablesOle.invoke(tableItmes[0], new Variant[]{new Variant(index)});
            OleAutomation currentTable = tableItmeVar.getAutomation();
            // 得到表格行数
            int rows[] = currentTable.getIDsOfNames(new String[]{"Rows"});
            Variant rowVar = currentTable.getProperty(rows[0]);
            OleAutomation rowOle = rowVar.getAutomation();
            //如果不是后缀模式，则删除之前的表格
            if (!append) {
                // 得到该表格的行数
                int[] rowCnt = rowOle.getIDsOfNames(new String[]{"Count"});
                Variant rowCntVar = rowOle.getProperty(rowCnt[0]);
                int rowsCnt = rowCntVar.getInt();
                // 遍历该表格所有单元格
                for (int i = 1; i < rowsCnt; i++) {
                    // 得到第2行的表格
                    int item[] = rowOle.getIDsOfNames(new String[]{"Item"});
                    Variant it = rowOle.invoke(item[0], new Variant[]{new Variant(2)});
                    currentTable = it.getAutomation();
                    // 删除第2行的表格
                    int deleteFunc[] = currentTable.getIDsOfNames(new String[]{"Delete"});
                    currentTable.invoke(deleteFunc[0], new Variant[]{});
                }
            }
            // 增加一行
            int addFunc[] = rowOle.getIDsOfNames(new String[]{"Add"});
            Variant currentRowVar = rowOle.invoke(addFunc[0], new Variant[]{});
            // 选中该行
            OleAutomation currentRowOle = currentRowVar.getAutomation();
            int select[] = currentRowOle.getIDsOfNames(new String[]{"Select"});
            currentRowOle.invoke(select[0]);
            // 遍历写入单元格值
            for (int i = 0; items != null && i < items.length; i++) {
                // 得到选中的所有的单元格
                int cells[] = currentRowOle.getIDsOfNames(new String[]{"Cells"});
                Variant crVar = currentRowOle.getProperty(cells[0]);
                OleAutomation crOle = crVar.getAutomation();
                // 得到要写入值的单元格
                int cellItmes[] = crOle.getIDsOfNames(new String[]{"Item"});
                Variant cellItmeVar = crOle.invoke(cellItmes[0], new Variant[]{new Variant(i + 1)});
                OleAutomation cellItmeVarOle = cellItmeVar.getAutomation();
                // 得到该单元格的范围
                int ranges[] = cellItmeVarOle.getIDsOfNames(new String[]{"Range"});
                Variant rangesVar = cellItmeVarOle.getProperty(ranges[0]);
                OleAutomation rangesOle = rangesVar.getAutomation();
                // 写入值
                int text1[] = rangesOle.getIDsOfNames(new String[]{"Text"});
                rangesOle.setProperty(text1[0], new Variant[]{new Variant(items[i])});
            }
            doc.dispose();
        } catch (org.eclipse.swt.SWTException e) {
            String str = "OleClientSite:" + e.toString();
            System.out.println(str);
            return;
        }
    }

    public void saveDoc() {
        site.save(new File(filePath), true);
    }

    public void saveDoc(String otherFilePath) {
        site.save(new File(otherFilePath), true);
    }

    //隐藏边框
    public void hideBorder() {
        doc = new OleAutomation(site);

        int selection[] = doc.getIDsOfNames(new String[]{"Parent"});
        Variant se = doc.getProperty(selection[0]);
        OleAutomation seOle = se.getAutomation();

        int addin[] = seOle.getIDsOfNames(new String[]{"AddIns"});
        Variant se2 = seOle.getProperty(addin[0]);
        OleAutomation seOle3 = se2.getAutomation();
        String url = WPSWrapper.class.getClassLoader().getResource("nc/wps/accessory/template.docx").getPath().substring(1);
        System.out.println(url);

        int addin1[] = seOle3.getIDsOfNames(new String[]{"Add"});
        Variant it = seOle3.invoke(addin1[0], new Variant[]{new Variant(url)});
    }

    /**
     * 打印文件对话框
     */
    public void printDialog() {
        doc = new OleAutomation(site);

        int selection[] = doc.getIDsOfNames(new String[]{"Parent"});
        Variant se = doc.getProperty(selection[0]);
        OleAutomation seOle = se.getAutomation();

        int dialogs[] = seOle.getIDsOfNames(new String[]{"Dialogs"});
        Variant bmk = seOle.getProperty(dialogs[0]);
        OleAutomation bmkBook = bmk.getAutomation();


        int item[] = bmkBook.getIDsOfNames(new String[]{"Item"});
        Variant it = bmkBook.invoke(item[0], new Variant[]{new Variant(88)});
        OleAutomation itAu = it.getAutomation();

        int item1[] = itAu.getIDsOfNames(new String[]{"Show"});
        it = itAu.invoke(item1[0]);
    }

    /**
     * 仅文本域可编辑
     */
    public void lock() {
        doc = new OleAutomation(site);
        int protect[] = doc.getIDsOfNames(new String[]{"Protect"});
        doc.invoke(protect[0], new Variant[]{new Variant(2), new Variant(true)});

    }

}
