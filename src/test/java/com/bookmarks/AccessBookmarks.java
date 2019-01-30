package com.bookmarks;

import com.aspose.words.Bookmark;
import com.aspose.words.Document;
import com.help.TestFile;
import org.junit.Test;

/**
 * @Auther: zch
 * @Date: 2019/1/30 14:29
 * @Description:获取一个书签
 */
public class AccessBookmarks {

    @Test
    public void test()throws Exception{
        String dataDir =  TestFile.getTestDataParentDir(this.getClass());
        Document doc = new Document(dataDir + "MyBookmark.docx");
        Bookmark bookmark1 = doc.getRange().getBookmarks().get(0);

        Bookmark bookmark = doc.getRange().getBookmarks().get("MyBookmark");
        doc.save(dataDir + "output.doc");
        System.out.println("\nTable bookmarked successfully.\nFile saved at " + dataDir);
    }

}
