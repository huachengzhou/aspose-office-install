package com.featurescomparison.workingwithimages;

import com.aspose.words.*;
import com.help.TestFile;
import org.testng.annotations.Test;

/**
 * @Auther: zch
 * @Date: 2019/1/11 17:56
 * @Description:
 */
public class AsposeInsertImage {

    @Test
    public void test() throws Exception {
        final String dataPath = TestFile.getTestDataParentDir(this.getClass()) + "data\\";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertImage(dataPath + "background.jpg");
        builder.insertImage(dataPath + "background.jpg",
                RelativeHorizontalPosition.MARGIN,
                100,
                RelativeVerticalPosition.MARGIN,
                200,
                200,
                100,
                WrapType.SQUARE);

        doc.save(dataPath + "Aspose_InsertImage_Out.docx");

        System.out.println("Process Completed Successfully");
    }

}
