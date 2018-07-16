package freemarker;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.HashMap;
import java.util.Map;

import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;

public class OutputText {

	public static void main(String[] args) throws IOException, TemplateException {
		// Ҫ��������, ע��map��keyҪ��word��${xxx}��xxxһ��
		Map<String, String> dataMap = new HashMap<String, String>();
		dataMap.put("username", "����");
		dataMap.put("sex", "��");

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
		File outFile = new File("C:\\\\Users\\\\Alan\\\\Desktop\\test_a.doc");

		// ��utf-8�ı����ȡftl�ļ�
		Template t = configuration.getTemplate("a.ftl", "utf-8");
		Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile), "utf-8"), 10240);
		t.process(dataMap, out);
		out.close();
		System.out.println("End");
	}

}
