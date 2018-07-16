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
		// Ҫ��������, ע��map��keyҪ��word��${xxx}��xxxһ��
		Map<String, Object> dataMap = new HashMap<String, Object>();
		dataMap.put("username", "����");
		dataMap.put("sex", "��");
		dataMap.put("imgStr", getImageStr());

		// Configuration���ڶ�ȡftl�ļ�
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
		 * ����������ָ��ftl�ļ�����Ŀ¼·���ķ�ʽ, ע�������ַ�ʽ���� ָ��ftl�ļ�����Ŀ¼��·��,������ftl�ļ���·��
		 */
		// ָ��·���ĵ�һ�ַ�ʽ(����ĳ��������·��ָ��)
		// configuration.setClassForTemplateLoading(this.getClass(),"");

		// ָ��·���ĵڶ��ַ�ʽ,�ҵ�·����C:/a.ftl
		configuration.setDirectoryForTemplateLoading(new File("C:\\Users\\Alan\\Desktop"));

		// ����ĵ�·��������
		File outFile = new File("C:\\Users\\Alan\\Desktop\\test_d.doc");

		// ��utf-8�ı����ȡftl�ļ�
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
