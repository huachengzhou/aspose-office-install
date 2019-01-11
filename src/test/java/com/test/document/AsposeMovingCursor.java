package com.test.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Node;
import com.aspose.words.Paragraph;
import com.help.Utils;
import org.testng.annotations.Test;

/**
 * @Auther: zch
 * @Date: 2019/1/11 12:02
 * @Description:word内容移动(类似与copy)
 */
public class AsposeMovingCursor {

    /**
     * word内容移动(类似与copy)
     * @throws Exception
     */
    @Test
    public void test()throws Exception{
        String dataPath = Utils.getDataDir(this.getClass());
        Document doc = new Document(dataPath + "房屋及公用设施等移交验收交接表.dotx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        //Shows how to access the current node in a document builder.
        Node curNode = builder.getCurrentNode();
        Paragraph curParagraph = builder.getCurrentParagraph();

        // Shows how to move a cursor position to a specified node.
        builder.moveTo(doc.getFirstSection().getBody().getLastParagraph());

        // Shows how to move a cursor position to the beginning or end of a document.
        builder.moveToDocumentEnd();
        builder.writeln("This is the end of the document.");

        builder.moveToDocumentStart();
        builder.writeln("This is the beginning of the document.");

        doc.save(dataPath + "AsposeMovingCursor.doc");

        System.out.println("Done.");
    }

}
