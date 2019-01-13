package com.featurescomparison.workingwithranges.insertbeforeandafterranges;

import com.help.TestFile;
import org.apache.poi.hwpf.HWPFDocument;
import org.testng.annotations.Test;

import java.io.FileInputStream;

/**
 * @Author noatn
 * @Description
 * @createDate 2019/1/13
 **/
public class ApacheInsert {


    @Test
    public void test() throws Exception {
        final String dataPath = TestFile.getTestDataParentDir(this.getClass()) + "data\\";
        HWPFDocument doc = new HWPFDocument(new FileInputStream(
                dataPath + "document.doc"));

        doc.getRange().getSection(0).insertBefore("Apache Inserted THIS Text before the below section");

        String text = doc.getRange().text();

        System.out.println("Range: " + text);
    }

}
