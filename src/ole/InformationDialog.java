package ole;

import java.awt.Font;
import java.awt.Label;
import java.awt.Toolkit;

import javax.swing.ImageIcon;


@SuppressWarnings("serial")
public class InformationDialog extends java.awt.Dialog {
    private String infor;
    private String title;

    public InformationDialog(boolean model, String title,String infor) {
        super(null, model);
        this.infor = infor;
        this.title = title;
        initComponents();
        int x = (Toolkit.getDefaultToolkit().getScreenSize().width - getSize().width) / 2;
        int y = (Toolkit.getDefaultToolkit().getScreenSize().height - getSize().height) / 2;
        setLocation(x, y);
        this.setResizable(false);
        this.setVisible(true);
        this.setAlwaysOnTop(true);
        this.setTitle(title);
        String url = InformationDialog.class.getClassLoader().getResource("ole/logo.png").getFile();
        ImageIcon icon = new ImageIcon(url);
        this.setIconImage(icon.getImage());
    }

    private void initComponents() {

        label1 = new Label(infor,Label.CENTER);
        label1.setFont(new Font("宋体", Font.PLAIN, 17));
        addWindowListener(new java.awt.event.WindowAdapter() {
            @Override
            public void windowClosing(java.awt.event.WindowEvent evt) {
                closeDialog(evt);
            }
        });

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(this);
        this.setLayout(layout);
        layout.setHorizontalGroup(
                layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(layout.createSequentialGroup()
                                .addContainerGap()
                                .addComponent(label1, javax.swing.GroupLayout.DEFAULT_SIZE, 313, Short.MAX_VALUE)
                                .addContainerGap())
        );
        layout.setVerticalGroup(
                layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(layout.createSequentialGroup()
                                .addGap(27, 27, 27)
                                .addComponent(label1, javax.swing.GroupLayout.PREFERRED_SIZE, 50, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addContainerGap(21, Short.MAX_VALUE))
        );

        pack();
    }

    private void closeDialog(java.awt.event.WindowEvent evt) {//GEN-FIRST:event_closeDialog
        setVisible(false);
        dispose();
    }

    public static void main(String args[]) {
        InformationDialog dialog = new InformationDialog(true,"提示信息","加载中，请稍候……");
    }

    private Label label1;
}
