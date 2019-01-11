package com.test.document;

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import com.help.Utils;
import org.testng.annotations.Test;

/**
 * @Auther: zch
 * @Date: 2019/1/11 13:59
 * @Description:重新修订
 */
public class AsposeTrackChanges {

    /**
     * 重新修订
     * @throws Exception
     */
    @Test
    public void test()throws Exception{
        String dataPath = Utils.getDataDir(this.getClass());
        Document doc = new Document(dataPath +"trackDoc.doc");
        doc.acceptAllRevisions();
        doc.save(dataPath + "AsposeAcceptChanges.doc", SaveFormat.DOC);

        System.out.println("Done.");
    }

}
