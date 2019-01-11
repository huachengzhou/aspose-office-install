package com.featurescomparison.workingwithdocuments.getdocumentproperties;

import com.aspose.words.Document;
import com.aspose.words.DocumentProperty;
import com.aspose.words.FileFormatInfo;
import com.aspose.words.FileFormatUtil;
import com.help.TestFile;
import org.testng.annotations.Test;

import java.io.File;
import java.text.MessageFormat;

/**
 * @Auther: zch
 * @Date: 2019/1/11 17:33
 * @Description:
 */
public class AsposeDocumentProperties {

    /**
     * 获取属性
     * @throws Exception
     */
    @Test
    public void test() throws Exception {
        String dataPath = new File(TestFile.getTestDataDir(this.getClass())).getParent();
        Document doc = new Document(dataPath + "\\document.doc");

        System.out.println("============ Built-in Properties ============");
        for (DocumentProperty prop : doc.getBuiltInDocumentProperties())
            System.out.println(MessageFormat.format("{0} : {1}", prop.getName(), prop.getValue()));

        System.out.println("============ Custom Properties ============");
        for (DocumentProperty prop : doc.getCustomDocumentProperties())
            System.out.println(MessageFormat.format("{0} : {1}", prop.getName(), prop.getValue()));

        FileFormatInfo info = FileFormatUtil.detectFileFormat(dataPath + "\\document.doc");
        System.out.println("The document format is: " + FileFormatUtil.loadFormatToExtension(info.getLoadFormat()));
        System.out.println("Document is encrypted: " + info.isEncrypted());
        System.out.println("Document has a digital signature: " + info.hasDigitalSignature());
    }

}
