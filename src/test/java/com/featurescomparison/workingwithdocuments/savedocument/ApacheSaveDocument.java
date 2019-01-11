package com.featurescomparison.workingwithdocuments.savedocument;

import com.help.TestFile;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.testng.annotations.Test;

import java.io.FileOutputStream;

/**
 * @Auther: zch
 * @Date: 2019/1/11 17:47
 * @Description:
 */
public class ApacheSaveDocument {

    @Test
    public void test() throws Exception {
        final String dataPath = TestFile.getTestDataParentDir(this.getClass());
        XWPFDocument document = new XWPFDocument();
        XWPFParagraph tmpParagraph = document.createParagraph();

        XWPFRun tmpRun = tmpParagraph.createRun();
        tmpRun.setText("Apache Sample Content for Word file.");

        FileOutputStream fos = new FileOutputStream(dataPath + "Apache_SaveDoc_Out.doc");
        document.write(fos);
        fos.close();

        System.out.println("Process Completed Successfully");
    }

}
