package ole;

import java.awt.*;

import javax.swing.*;

/**
 * @author Alan
 */
public class S {
    JFrame frame = new JFrame();

    S(String info) {
//        dialog.setSize(600, 200);

        Container contentPane = frame.getContentPane();
        contentPane.setLayout(new BorderLayout());
//        contentPane.setLayout(null);

//        setLayout(new BorderLayout());

        JLabel label = new JLabel(info);
//        label.setSize(150, 50);
        label.setFont(new Font("宋体", Font.PLAIN, 25));
//        add(label);
//        add(label, BorderLayout.CENTER);

//        JPanel panel = new JPanel();
//        panel.setSize(300, 100);
//        panel.setLayout(new BorderLayout());
//        panel.add(label, BorderLayout.CENTER);

        contentPane.add(label, BorderLayout.CENTER);
//        contentPane.add(label);

//        add(panel);

//        String url = InformationDialog.class.getClassLoader().getResource("ole/logo.png").getFile();
//        ImageIcon icon = new ImageIcon(url);
//        setIconImage(icon.getImage());
//        setUndecorated(true);

//        dialog.setResizable(false);
        frame.setAlwaysOnTop(true);
        frame.setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);

        frame.pack();

        frame.setSize(600, 200);

        int x = (Toolkit.getDefaultToolkit().getScreenSize().width - 600) / 2;
        int y = (Toolkit.getDefaultToolkit().getScreenSize().height - 200) / 2;
        frame.setLocation(x, y);
    }

    void show() {
        frame.setVisible(true);
    }

    void dispose() {
        frame.dispose();
    }

    public static void main(String args[]) throws InterruptedException {
        S s = new S("你好");
        s.show();
        Thread.sleep(3000);
//        s.dispose();
    }

}
