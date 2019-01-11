package com.featurescomparison.workingwithdocuments.getdocumentproperties;

import com.help.TestFile;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hwpf.HWPFDocument;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;

/**
 * @Auther: zch
 * @Date: 2019/1/11 17:29
 * @Description:
 */
public class ApacheDocumentProperties {

    /**
     * 获取属性
     * @throws Exception
     */
    @Test
    public void test() throws Exception {
        String dataPath = new File(TestFile.getTestDataDir(this.getClass())).getParent();
        HWPFDocument doc = new HWPFDocument(new FileInputStream(
                dataPath + "\\document.doc"));
        SummaryInformation summaryInfo = doc.getSummaryInformation();
        System.out.println(summaryInfo.getApplicationName());
        System.out.println(summaryInfo.getAuthor());
        System.out.println(summaryInfo.getComments());
        System.out.println(summaryInfo.getCharCount());
        System.out.println(summaryInfo.getEditTime());
        System.out.println(summaryInfo.getKeywords());
        System.out.println(summaryInfo.getLastAuthor());
        System.out.println(summaryInfo.getPageCount());
        System.out.println(summaryInfo.getRevNumber());
        System.out.println(summaryInfo.getSecurity());
        System.out.println(summaryInfo.getSubject());
        System.out.println(summaryInfo.getTemplate());
    }
}
