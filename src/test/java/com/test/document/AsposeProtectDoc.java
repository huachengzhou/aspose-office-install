package com.test.document;

import com.aspose.words.*;
import com.help.Utils;
import org.testng.annotations.Test;

/**
 * @Auther: zch
 * @Date: 2019/1/11 13:50
 * @Description:设置文档只读属性
 */
public class AsposeProtectDoc {

    /**
     * 设置文档只读属性
     * @throws Exception
     */
    @Test
    public void test()throws Exception{
        String dataPath = Utils.getDataDir(this.getClass());
        Document doc = new Document(dataPath + "客户数.dotx");
        doc.protect(ProtectionType.READ_ONLY);
//		doc.protect(ProtectionType.ALLOW_ONLY_COMMENTS);
//		doc.protect(ProtectionType.ALLOW_ONLY_FORM_FIELDS);
//		doc.protect(ProtectionType.ALLOW_ONLY_REVISIONS);

        doc.save(dataPath + "AsposeProtect.doc", SaveFormat.DOC);
        System.out.println("Done.");
    }

}
