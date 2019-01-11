package com.featurescomparison.workingwithdocuments.createnewdocument;

import com.help.TestFile;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.testng.annotations.Test;

import java.io.FileOutputStream;

/**
 * @Auther: zch
 * @Date: 2019/1/11 17:12
 * @Description:apache 方式创建word
 */
public class ApacheNewDocument {

    /**
     * apache 方式创建word
     * @throws Exception
     */
    @Test
    public void test()throws Exception{
        final String dataPath = TestFile.getTestDataDir(this.getClass());
        XWPFDocument document = new XWPFDocument();
        XWPFParagraph tmpParagraph = document.createParagraph();
        XWPFRun tmpRun = tmpParagraph.createRun();
        tmpRun.setText("Apache Sample Content for Word file.");
        tmpRun.setFontSize(18);

        FileOutputStream fos = new FileOutputStream(dataPath + "Apache_newWordDoc_Out.doc");
        document.write(fos);
        fos.close();
    }

}
