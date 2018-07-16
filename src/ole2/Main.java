package ole2;

import java.awt.Point;

public class Main {

	public static void main(String[] args) {
		String filePath = "C:\\Users\\Alan\\Desktop\\test.doc";
		WPSDialog wpsDialog = new WPSDialog(filePath, new Point(600,200), 670, 720, 2);
	}

}
