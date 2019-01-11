package com.featurescomparison.workingwithimages;

import com.help.TestFile;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;

/**
 * @Auther: zch
 * @Date: 2019/1/11 17:56
 * @Description:
 */
public class ApacheInsertImage {

    @Test
    public void test() throws Exception {
        final String dataPath = TestFile.getTestDataParentDir(this.getClass()) + "data\\";

        XWPFDocument doc = new XWPFDocument();
        XWPFParagraph p = doc.createParagraph();

        String imgFile = dataPath + "aspose.jpg";
        XWPFRun r = p.createRun();

        int format = XWPFDocument.PICTURE_TYPE_JPEG;
        r.setText(imgFile);
        r.addBreak();
        r.addPicture(new FileInputStream(imgFile), format, imgFile, Units.toEMU(200), Units.toEMU(200)); // 200x200 pixels
        r.addBreak(BreakType.PAGE);

        FileOutputStream out = new FileOutputStream(dataPath + "Apache_ImagesInDoc_Out.docx");
        doc.write(out);
        out.close();

        System.out.println("Process Completed Successfully");
    }

}
