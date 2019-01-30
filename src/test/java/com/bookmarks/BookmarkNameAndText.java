package com.bookmarks;

import com.aspose.words.Bookmark;
import com.aspose.words.Document;
import com.help.TestFile;
import org.testng.annotations.Test;

/**
 * @Auther: zch
 * @Date: 2019/1/30 14:35
 * @Description:书签名称和内容重新命名
 */
public class BookmarkNameAndText {

    @Test
    public void test()throws Exception{
        String dataDir =  TestFile.getTestDataParentDir(this.getClass());
        Document doc = new Document(dataDir + "MyBookmark.docx");
        // Use the indexer of the Bookmarks collection to obtain the desired bookmark.
        Bookmark bookmark = doc.getRange().getBookmarks().get("MyBookmark");

        // Get the name and text of the bookmark.
        String name = bookmark.getName();
        String text = bookmark.getText();

        // Set the name and text of the bookmark.
        bookmark.setName("RenamedBookmark");
        bookmark.setText("This is a new bookmarked text.");

        doc.save(dataDir + "output.doc");
        System.out.println("\nBookmark name and text set successfully.");

        //ExEnd:BookmarkNameAndText
    }

}
