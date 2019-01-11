package com.test.text;

import com.aspose.words.*;
import com.help.Utils;
import org.testng.annotations.Test;

import java.util.Date;

/**
 * @Auther: zch
 * @Date: 2019/1/11 15:11
 * @Description:插入批注
 */
public class AsposeInsertComments {

    /**
     * 插入批注
     * @throws Exception
     */
    @Test
    public void test() throws Exception{
        final String dataPath = Utils.getDataDir(this.getClass());
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("Some text is added.");

        Comment comment = new Comment(doc, "Aspose", "As", new Date());
        builder.getCurrentParagraph().appendChild(comment);
        comment.getParagraphs().add(new Paragraph(doc));
        comment.getFirstParagraph().getRuns().add(new Run(doc, "Comment text."));

        doc.save(dataPath + "AsposeComments.docx");

        System.out.println("Done.");
    }

}
