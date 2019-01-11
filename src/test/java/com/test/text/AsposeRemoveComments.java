package com.test.text;

import com.aspose.words.Comment;
import com.aspose.words.Document;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.help.Utils;
import org.testng.annotations.Test;

/**
 * @Auther: zch
 * @Date: 2019/1/11 15:24
 * @Description:删除批注
 */
public class AsposeRemoveComments {
    /**
     * 删除批注
     *
     * @throws Exception
     */
    @Test
    public void test() throws Exception {
        final String dataPath = Utils.getDataDir(this.getClass());
        Document doc = new Document(dataPath + "AsposeComments.docx");

        // Collect all comments in the document
        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);
        // Look through all comments and remove those written by the authorName author.
        for (int i = comments.getCount() - 1; i >= 0; i--) {
            Comment comment = (Comment) comments.get(i);
            if (comment.getAuthor().equalsIgnoreCase("Aspose"))
                System.out.println("Aspose comment removed");
            comment.remove();
        }

        doc.save(dataPath + "AsposeCommentsRemoved.docx");
        System.out.println("Done...");
    }
}
