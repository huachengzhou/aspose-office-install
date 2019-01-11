package com.test.text;

import com.aspose.words.*;
import com.help.Utils;
import org.testng.annotations.Test;

import java.util.ArrayList;

/**
 * @Auther: zch
 * @Date: 2019/1/11 15:19
 * @Description:获取批注
 */
public class AsposeExtractComments {
    /**
     * 获取批注
     * @throws Exception
     */
    @Test
    public void test() throws Exception {
        final String dataPath = Utils.getDataDir(this.getClass());
        Document doc = new Document(dataPath + "AsposeComments.docx");

        ArrayList collectedComments = new ArrayList();

        // Collect all comments in the document
        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);

        // Look through all comments and gather information about them.
        for (Comment comment : (Iterable<Comment>) comments) {
            System.out.println(comment.getAuthor() + " - " + comment.getDateTime() + " - "
                    + comment.toString(SaveFormat.TEXT));
        }
        System.out.println("Done.");
    }
}
