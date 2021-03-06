package ole;

import java.io.IOException;
import java.io.InputStream;
import java.io.ObjectInputStream;
import java.util.ArrayList;
import java.util.HashMap;

public class Main {

    public static void main(String[] args) throws IOException, ClassNotFoundException {
        InputStream is = Main.class.getResourceAsStream("测试君 20180604-161524");
        ObjectInputStream ois = new ObjectInputStream(is);
        HashMap<String,String> headValueMap = (HashMap<String, String>) ois.readObject();
        HashMap<Integer, ArrayList<String[]>> bodyValueMap = (HashMap<Integer, ArrayList<String[]>>) ois.readObject();
        System.out.println(headValueMap);
        System.out.println(bodyValueMap);
        String photoPath = Main.class.getClassLoader().getResource("ole/logo.png").getFile();
        photoPath = photoPath.substring(1, photoPath.length());
        System.out.println(photoPath);
        String filePath = Main.class.getClassLoader().getResource("ole/abc.doc").getFile();
        System.out.println(filePath);
        new WPSAccessoryEditor(filePath, true, 1);
    }

}
