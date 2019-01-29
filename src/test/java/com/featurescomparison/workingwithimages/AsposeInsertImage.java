package com.featurescomparison.workingwithimages;

import com.aspose.words.*;
import com.help.TestFile;
import org.apache.tools.ant.taskdefs.Exec;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.Test;

import java.io.File;

/**
 * @Auther: zch
 * @Date: 2019/1/11 17:56
 * @Description:
 */
public class AsposeInsertImage {

    private final Logger log = LoggerFactory.getLogger(getClass());

    @Test
    public void test() throws Exception {
        final String dataPath = TestFile.getTestDataParentDir(this.getClass()) + "\\data\\test.docx";
        System.out.println(dataPath);

        Document doc = new Document(dataPath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        Bookmark bookmark = doc.getRange().getBookmarks().get("test");
        System.out.println("bookmark:"+bookmark.getText());

        String imgPath = TestFile.getTestDataParentDir(this.getClass())+"\\data\\" + "background.jpg" ;
        System.out.println(imgPath);
        builder.moveToBookmark("test");
        builder.insertImage(imgPath,
                RelativeHorizontalPosition.DEFAULT,
                100,
                RelativeVerticalPosition.MARGIN,
                200,
                200,
                100,
                WrapType.SQUARE);

        doc.save(TestFile.getTestDataParentDir(this.getClass())+"data\\" + "Aspose_InsertImage_Out.docx");

        System.out.println("Process Completed Successfully");
    }

    public String exportWmfFromDoc(String fileName) {
       return null;
    }

    public void test2() throws Exception {
        String templatePath = "";
        Document doc = new Document(templatePath);
        for (Bookmark mark : doc.getRange().getBookmarks()) {
            if (mark != null) {
                switch (mark.getName()) {
                    case "NAME":
                        mark.setText("龚辉");
                        break;
                    case "PHOTO":
                        DocumentBuilder builder = new DocumentBuilder(doc);
                        String imgPath = "";
                        builder.moveToBookmark("PHOTO");
                        builder.insertImage(imgPath + "background.jpg",
                                RelativeHorizontalPosition.MARGIN,
                                100,
                                RelativeVerticalPosition.MARGIN,
                                200,
                                200,
                                100,
                                WrapType.SQUARE);
                        break;
                    default:
                        break;
                }
            }
        }

    }

}
