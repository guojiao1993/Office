package ole;

import java.awt.*;

import javax.swing.*;

/**
 * @author Alan
 */
public class InformationDialog extends JDialog {

    InformationDialog(String info) {
        Label label = new Label(info, Label.CENTER);
        label.setFont(new Font("宋体", Font.PLAIN, 17));

        setLayout(new BorderLayout());
        add(label, BorderLayout.CENTER);

        int x = (Toolkit.getDefaultToolkit().getScreenSize().width - getSize().width) / 2;
        int y = (Toolkit.getDefaultToolkit().getScreenSize().height - getSize().height) / 2;
        setLocation(x, y);

        String url = InformationDialog.class.getClassLoader().getResource("ole/logo.png").getFile();
        ImageIcon icon = new ImageIcon(url);
        setIconImage(icon.getImage());

        setResizable(false);
        setAlwaysOnTop(true);
        setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);

        pack();
        setVisible(true);
    }

    public static void main(String args[]) throws InterruptedException {
        InformationDialog dialog = new InformationDialog("你好");
        Thread.sleep(3000);
        dialog.dispose();
    }

}
