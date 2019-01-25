package org.test;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.help.TestFile;
import org.testng.annotations.Test;

/**
 * @Auther: zch
 * @Date: 2019/1/25 15:32
 * @Description:
 */
public class Demo4 {

    @Test
    public void replaceTest()throws Exception{
        final String dataPath = TestFile.getTestDataParentDir(this.getClass());
        Document doc = new Document(String.format("%s%s",dataPath,"房产预评报告模板.doc"));
        DocumentBuilder builder = new DocumentBuilder(doc);

        doc.getRange().replace("gags","asgsag",false,false);

//        doc.save(String.format("%s%s%s%s", dataPath, this.getClass().getSimpleName(), Thread.currentThread().getStackTrace()[1].getMethodName(), ".docx"));
    }

}
