package com.test.document;

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import com.help.Utils;
import org.testng.annotations.Test;

/**
 * @Auther: zch
 * @Date: 2019/1/11 11:36
 * @Description:word复制
 */
public class AsposeCloneDoc {

    /**
     * word复制
     * @throws Exception
     */
    @Test
    public void test()throws Exception{
        String dataDir = Utils.getDataDir(this.getClass());

        Document doc = new Document(dataDir + "产品报价表模板.doc");
        Document clone = doc.deepClone();
        clone.save(dataDir + "AsposeClone.doc", SaveFormat.DOC);

        System.out.println("Done.");
    }

}
