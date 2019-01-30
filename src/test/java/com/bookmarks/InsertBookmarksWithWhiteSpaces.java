package com.bookmarks;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.PdfSaveOptions;
import com.help.TestFile;
import org.testng.annotations.Test;

/**
 * @Auther: zch
 * @Date: 2019/1/30 14:58
 * @Description:创建pdf书签以及内容(主从关系)
 */
public class InsertBookmarksWithWhiteSpaces {


    @Test
    public void test()throws Exception{
        String dataDir =  TestFile.getTestDataParentDir(this.getClass());
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startBookmark("My Bookmark");
        builder.writeln("Text inside a bookmark.");

        builder.startBookmark("Nested Bookmark");
        builder.writeln("Text inside a NestedBookmark.");
        builder.endBookmark("Nested Bookmark");

        builder.writeln("Text after Nested Bookmark.");
        builder.endBookmark("My Bookmark");


        PdfSaveOptions options = new PdfSaveOptions();
        options.getOutlineOptions().getBookmarksOutlineLevels().add("My Bookmark", 1);
        options.getOutlineOptions().getBookmarksOutlineLevels().add("Nested Bookmark", 2);

        dataDir = dataDir + "Insert.Bookmarks_out_.pdf";
        doc.save(dataDir, options);

        System.out.println("\nBookmarks with white spaces inserted successfully.\nFile saved at " + dataDir);
        //ExEnd:InsertBookmarksWithWhiteSpaces
    }

}
