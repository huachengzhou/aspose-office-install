package com.test.document;

import com.aspose.words.Document;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.SaveFormat;
import com.help.Utils;
import org.testng.annotations.Test;

/**
 * @Auther: zch
 * @Date: 2019/1/11 11:06
 * @Description:文档合并
 */
public class AppendDoc {

    /**
     * 两个文档合并
     * @throws Exception
     */
    @Test
    public void test() throws Exception {
        String dataDir = Utils.getDataDir(this.getClass());
        System.out.println(dataDir);
        Document doc1 = new Document(dataDir + "doc1.doc");
        Document doc2 = new Document(dataDir + "doc2.doc");

        doc1.appendDocument(doc2, ImportFormatMode.KEEP_SOURCE_FORMATTING);

        doc1.save(dataDir + "AsposeMerged.doc", SaveFormat.DOC);
    }
}
