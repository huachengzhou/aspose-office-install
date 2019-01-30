package com.bookmarks;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.PdfSaveOptions;
import com.help.TestFile;
import org.testng.annotations.Test;

/**
 * @Auther: zch
 * @Date: 2019/1/30 14:51
 * @Description:创建pdf书签以及内容
 */
public class CreateBookmark {

    @Test
    public void test()throws Exception{
        String dataDir =  TestFile.getTestDataParentDir(this.getClass());
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startBookmark("MyBookmark");
        builder.writeln("Text inside a bookmark.");
        builder.endBookmark("MyBookmark");

        builder.startBookmark("Nested Bookmark");
        builder.writeln("Text inside a NestedBookmark.");
        builder.endBookmark("Nested Bookmark");

        builder.writeln("Text after Nested Bookmark.");
        builder.endBookmark("My Bookmark");

        PdfSaveOptions options = new PdfSaveOptions();
        options.getOutlineOptions().setDefaultBookmarksOutlineLevel(1);
        options.getOutlineOptions().setDefaultBookmarksOutlineLevel(2);

        doc.save(dataDir + "output.pdf", options);
        System.out.println("\nBookmark created successfully.");
        //ExEnd:CreateBookmark
    }

}
