package com.test.document;

import com.aspose.words.Bookmark;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.SaveFormat;
import com.help.Utils;
import org.testng.annotations.Test;

/**
 * @Auther: zch
 * @Date: 2019/1/11 14:05
 * @Description:
 */
public class AsposeBookmarks {

    /**
     * 书签设置
     * @throws Exception
     */
    @Test
    public void test()throws Exception{
        String dataPath = Utils.getDataDir(this.getClass());

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startBookmark("AsposeBookmark");
        builder.writeln("Text inside a bookmark.");
        builder.endBookmark("AsposeBookmark");

        // By index.
        Bookmark bookmark1 = doc.getRange().getBookmarks().get(0);

        // By name.
        Bookmark bookmark2 = doc.getRange().getBookmarks().get("AsposeBookmark");

        doc.save(dataPath + "AsposeBookmark.doc", SaveFormat.DOC);

        System.out.println("Done.");

    }
}
