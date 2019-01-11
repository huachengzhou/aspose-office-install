package com.featurescomparison.workingwithdocuments.formattext;

import com.help.TestFile;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.testng.annotations.Test;

import java.io.FileOutputStream;

/**
 * @Auther: zch
 * @Date: 2019/1/11 17:21
 * @Description:格式化
 */
public class ApacheFormattedText {

    /**
     * 格式化
     * @throws Exception
     */
    @Test
    public void test()throws Exception{
        final String dataPath = TestFile.getTestDataDir(this.getClass());
        // Create a new document from scratch
        XWPFDocument doc = new XWPFDocument();

        // create paragraph
        XWPFParagraph para = doc.createParagraph();

        // create a run to contain the content
        XWPFRun rh = para.createRun();

        // Format as desired
        rh.setFontSize(15);
        rh.setFontFamily("Verdana");
        rh.setText("This is the formatted Text");
        rh.setColor("fff000");
        para.setAlignment(ParagraphAlignment.RIGHT);

        // write the file
        FileOutputStream out = new FileOutputStream(dataPath + "Apache_FormattedText_Out.docx");
        doc.write(out);
        out.close();

        System.out.println("Process Completed Successfully");
    }

}
