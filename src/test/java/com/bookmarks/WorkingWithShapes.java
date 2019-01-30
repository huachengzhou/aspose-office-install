package com.bookmarks;

import com.aspose.words.*;
import com.help.TestFile;
import org.testng.annotations.Test;

/**
 * @Auther: zch
 * @Date: 2019/1/30 17:11
 * @Description:
 */
public class WorkingWithShapes {
    private  String dataDir =  TestFile.getTestDataParentDir(this.getClass());

    /**
     * 插入图片
     * @throws Exception
     */
    @Test
    public void insertShapeUsingDocumentBuilder()throws Exception{
        // ExStart:InsertShapeUsingDocumentBuilder
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        //Free-floating shape insertion. 自由浮动形状插入。
        Shape shape = builder.insertImage(dataDir + "Images\\dotnet-logo.png",
                RelativeHorizontalPosition.PAGE, 100,
                RelativeVerticalPosition.PAGE, 0,
                50, 50,
                WrapType.NONE);

        shape.setRotation(30.0);

        builder.writeln();

        //Inline shape insertion. 内嵌形状插入。
        shape =  builder.insertImage(dataDir + "Images\\Test_636_852.gif",
                RelativeHorizontalPosition.PAGE, 100,
                RelativeVerticalPosition.PAGE, 100,
                50, 50,
                WrapType.NONE);
        shape.setRotation(30.0);

        OoxmlSaveOptions so = new OoxmlSaveOptions(SaveFormat.DOCX);
        // "Strict" or "Transitional" compliance allows to save shape as DML.
        so.setCompliance(OoxmlCompliance.ISO_29500_2008_TRANSITIONAL);

        dataDir = dataDir + "Shape_InsertShapeUsingDocumentBuilder_out.docx";

        // Save the document to disk.
        doc.save(dataDir, so);
        // ExEnd:InsertShapeUsingDocumentBuilder
        System.out.println("\nInsert Shape successfully using DocumentBuilder.\nFile saved at " + dataDir);
    }

}
