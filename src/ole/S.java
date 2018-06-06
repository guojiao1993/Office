package ole;

import com.sun.awt.AWTUtilities;

import java.awt.*;

import javax.swing.*;

/**
 * @author Alan
 */
public class S extends JDialog{

    S(String info) {
        Container contentPane = getContentPane();
        contentPane.setLayout(new BorderLayout());

//        JLabel label = new JLabel(info);
//        label.setFont(new Font("宋体", Font.PLAIN, 25));
//        label.setHorizontalAlignment(SwingConstants.CENTER);
//        contentPane.add(label, BorderLayout.NORTH);

        String url = InformationDialog.class.getClassLoader().getResource("ole/bg.png").getFile();
        ImageIcon icon = new ImageIcon(url);
        setIconImage(icon.getImage());

        JLabel backLabel = new JLabel() {
            @Override
            public void paint(Graphics g) {
                super.paint(g);
                icon.paintIcon(this, g, 0, 0);
            }
        };
        add(backLabel, BorderLayout.CENTER);

        setUndecorated(true);
        setResizable(false);
        setAlwaysOnTop(true);
        setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);

        pack();

        setSize(300, 100);
//        int x = (Toolkit.getDefaultToolkit().getScreenSize().width - getWidth()) / 2;
//        int y = (Toolkit.getDefaultToolkit().getScreenSize().height - getHeight()) / 2;
//        setLocation(x, y);

        AWTUtilities.setWindowOpaque(this, false);
        setLocationRelativeTo(null);  //设置窗口居中

        setVisible(true);
    }

    public static void main(String args[]) throws InterruptedException {
        S s = new S("记载中，请稍候……");
        Thread.sleep(8000);
        s.dispose();
    }

}
