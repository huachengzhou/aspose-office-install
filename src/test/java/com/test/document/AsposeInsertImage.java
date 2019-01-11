package com.test.document;

import com.aspose.words.*;
import com.help.Utils;
import org.testng.annotations.Test;

/**
 * @Auther: zch
 * @Date: 2019/1/11 11:44
 * @Description:图片插入
 */
public class AsposeInsertImage {

    /**
     * 图片插入
     * @throws Exception
     */
    @Test
    public void test()throws Exception{
        String dataDir = Utils.getDataDir(this.getClass());

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertImage(dataDir + "background.jpg");
        builder.insertImage(dataDir + "background.jpg",
                RelativeHorizontalPosition.MARGIN,
                100,
                RelativeVerticalPosition.MARGIN,
                200,
                200,
                100,
                WrapType.SQUARE);

        doc.save(dataDir + "AsposeImageInDoc.docx");

        System.out.println("insert image Done.");
    }
}
