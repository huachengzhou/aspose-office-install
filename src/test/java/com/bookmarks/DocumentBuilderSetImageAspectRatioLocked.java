package com.bookmarks;

import com.aspose.words.*;
import com.help.TestFile;
import org.testng.annotations.Test;

/**
 * @Auther: zch
 * @Date: 2019/1/30 15:16
 * @Description:
 */
public class DocumentBuilderSetImageAspectRatioLocked {


    @Test
    public void test3()throws Exception{

    }

    /**
     * 显示如何在页面中间插入浮动图像。
     * @throws Exception
     */
    @Test
    public void test2() throws Exception {
        String dataDir = TestFile.getTestDataParentDir(this.getClass());
        // This creates a builder and also an empty document inside the builder.
        DocumentBuilder builder = new DocumentBuilder();

        // By default, the image is inline.
        Shape shape = builder.insertImage(dataDir + "Images\\Aspose.Words.gif");

        // Make the image float, put it behind text and center on the page.
        shape.setWrapType(WrapType.NONE);
        shape.setBehindText(true);
        shape.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
        shape.setHorizontalAlignment(HorizontalAlignment.CENTER);
        shape.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);
        shape.setVerticalAlignment(VerticalAlignment.CENTER);

        builder.getDocument().save(dataDir + "Image.CreateFloatingPageCenterOut.doc");
    }

    @Test
    public void test() throws Exception {
        //ExStart:DocumentBuilderSetImageAspectRatioLocked
        // The path to the documents directory.
        String dataDir = TestFile.getTestDataParentDir(this.getClass());

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape shape = builder.insertImage(dataDir + "Test.png");
//        shape.setAspectRatioLocked(false);
        shape.setAnchorLocked(false);

        doc.save(dataDir + "output.doc");
        //ExEnd:DocumentBuilderSetImageAspectRatioLocked
    }

    /**
     *
     * 表示绘图层中的对象，例如自选图形、文本框、自由形式、OLE对象、ActiveX控件或图片。
     使用形状类，可以在Microsoft Word文档中创建或修改形状。
     形状的一个重要属性是它的shapeType。不同类型的形状在Word文档中具有不同的功能。例如，只有图像和OLE形状可以在其中包含图像。大多数形状都可以有文本，但不是全部。
     形状可以包含文本，也可以包含段落和表节点作为子级。
     * */

}
