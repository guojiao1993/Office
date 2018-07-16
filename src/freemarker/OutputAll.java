package freemarker;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import sun.misc.BASE64Encoder;

public class OutputAll {

	public static void main(String[] args) throws IOException, TemplateException {
		// 要填充的数据, 注意map的key要和word中${xxx}的xxx一致
		Map<String, Object> dataMap = new HashMap<String, Object>();
		dataMap.put("username", "张三");
		dataMap.put("sex", "男");
		dataMap.put("imgStr", getImageStr());

		// Configuration用于读取ftl文件
		Configuration configuration = new Configuration();
		configuration.setDefaultEncoding("utf-8");

		ArrayList<Map<String, String>> list = new ArrayList<Map<String, String>>();
		for (int i = 0; i < 3; i++) {
			Map<String, String> temp = new HashMap<String, String>();
			temp.put("a", "a" + i);
			temp.put("b", "b" + i);
			temp.put("c", "c" + i);
			list.add(temp);
		}
		dataMap.put("userList", list);
		
		/*
		 * 以下是两种指定ftl文件所在目录路径的方式, 注意这两种方式都是 指定ftl文件所在目录的路径,而不是ftl文件的路径
		 */
		// 指定路径的第一种方式(根据某个类的相对路径指定)
		// configuration.setClassForTemplateLoading(this.getClass(),"");

		// 指定路径的第二种方式,我的路径是C:/a.ftl
		configuration.setDirectoryForTemplateLoading(new File("C:\\Users\\Alan\\Desktop"));

		// 输出文档路径及名称
		File outFile = new File("C:\\Users\\Alan\\Desktop\\test_d.doc");

		// 以utf-8的编码读取ftl文件
		Template t = configuration.getTemplate("d.ftl", "utf-8");
		Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile), "utf-8"), 10240);
		t.process(dataMap, out);
		out.close();
		System.out.println("End");
	}
	
	public static String getImageStr() {  
        String imgFile = "C:\\Users\\Alan\\Desktop\\b.jpg";  
        InputStream in = null;  
        byte[] data = null;  
        try {  
            in = new FileInputStream(imgFile);  
            data = new byte[in.available()];  
            in.read(data);  
            in.close();  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
        BASE64Encoder encoder = new BASE64Encoder();  
        return encoder.encode(data);  
    }
	
}
