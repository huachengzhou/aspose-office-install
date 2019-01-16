package com.test.fields;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Field;
import com.help.TestFile;
import com.help.Utils;
import org.testng.annotations.Test;

/**
 * @Auther: zch
 * @Date: 2019/1/11 15:06
 * @Description:字段删除
 */
public class AsposeRemoveFields {

    /**
     * 字段删除
     * @throws Exception
     */
    @Test
    public void test() throws Exception{
        final String dataPath = TestFile.getTestDataParentDir(this.getClass());


        Document doc = new Document();

        DocumentBuilder builder = new DocumentBuilder(doc);

        Field field = builder.insertField("PAGE");

        // Calling this method completely removes the field from the document.
        field.remove();

        doc.save(dataPath + "AsposeFieldsRemoved.docx");
        System.out.println("Aspose Fields Removed...");

    }

}
