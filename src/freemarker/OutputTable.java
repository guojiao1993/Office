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

public class OutputTable {

	public static void main(String[] args) throws IOException, TemplateException {
		new OutputTable().exportListWord();
	}

	public void exportListWord() throws IOException, TemplateException {
		// Ҫ��������, ע��map��keyҪ��word��${xxx}��xxxһ��
		HashMap<String, Object> dataMap = new HashMap<String, Object>();
		// dataMap.put("username", "����");
		// dataMap.put("sex", "��");
		// dataMap.put("imgStr", getImageStr());

		// ��������
//		ArrayList<User> list = new ArrayList<User>();
//		for (int i = 0; i < 10; i++) {
//			User user = new User("a" + (i + 1), "b" + (i + 1), "c" + (i + 1));
//			list.add(user);
//		}
//		dataMap.put("userList", list);
		
		ArrayList<Map<String, String>> list = new ArrayList<Map<String, String>>();
		for (int i = 0; i < 3; i++) {
			Map<String, String> temp = new HashMap<String, String>();
//			temp.put("a", " ");
			temp.put("b", "b" + i);
			temp.put("c", "c" + i);
			list.add(temp);
		}
		dataMap.put("userList", list);
		System.out.println(dataMap);

		// Configuration���ڶ�ȡftl�ļ�
		Configuration configuration = new Configuration();
		configuration.setDefaultEncoding("utf-8");

		/*
		 * ����������ָ��ftl�ļ�����Ŀ¼·���ķ�ʽ, ע�������ַ�ʽ���� ָ��ftl�ļ�����Ŀ¼��·��,������ftl�ļ���·��
		 */
		// ָ��·���ĵ�һ�ַ�ʽ(����ĳ��������·��ָ��)
		// configuration.setClassForTemplateLoading(this.getClass(),"");

		// ָ��·���ĵڶ��ַ�ʽ,�ҵ�·����C:/a.ftl
		configuration.setDirectoryForTemplateLoading(new File("C:\\Users\\Alan\\Desktop"));

		// ����ĵ�·��������
		File outFile = new File("C:\\Users\\Alan\\Desktop\\test_c.doc");

		// ��utf-8�ı����ȡftl�ļ�
		Template t = configuration.getTemplate("c.ftl", "utf-8");
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

	class User {
		public String a;
		public String b;
		public String c;

		public User(String a, String b, String c) {
			this.a = a;
			this.b = b;
			this.c = c;
		}

		public String getA() {
			return a;
		}

		public void setA(String a) {
			this.a = a;
		}

		public String getB() {
			return b;
		}

		public void setB(String b) {
			this.b = b;
		}

		public String getC() {
			return c;
		}

		public void setC(String c) {
			this.c = c;
		}

	}

}