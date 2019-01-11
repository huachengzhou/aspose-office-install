package com.featurescomparison.workingwithdocuments.converttoformats;

import com.help.TestFile;
import org.apache.poi.hwpf.HWPFDocumentCore;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.converter.WordToHtmlUtils;
import org.testng.annotations.Test;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

/**
 * @Auther: zch
 * @Date: 2019/1/11 16:08
 * @Description:
 */
public class ApacheConvertToFormats {

    /**
     * apache 方式将word转为html
     * @throws Exception
     */
    @Test
    public void test()throws Exception{
        final String dataPath = TestFile.getTestDataDir(this.getClass());
        File file  = new File(dataPath);
        if (!file.isDirectory()){
            file.mkdir();
        }
        HWPFDocumentCore wordDocument = WordToHtmlUtils.loadDoc(new FileInputStream(file.getParent() + "\\document.doc"));
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
                DocumentBuilderFactory.newInstance().newDocumentBuilder()
                        .newDocument());
        wordToHtmlConverter.processDocument(wordDocument);
        Document htmlDocument = wordToHtmlConverter.getDocument();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        DOMSource domSource = new DOMSource(htmlDocument);
        StreamResult streamResult = new StreamResult(out);

        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer serializer = tf.newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");
        serializer.transform(domSource, streamResult);
        out.close();

        FileOutputStream outputStream = new FileOutputStream(dataPath + "Apache_DocToHTML_Out.html");
        outputStream.write(out.toByteArray());
        outputStream.close();

        System.out.println("Apache - Doc file converted in specified formats");
    }

}
