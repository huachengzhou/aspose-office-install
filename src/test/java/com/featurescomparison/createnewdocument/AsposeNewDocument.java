package com.featurescomparison.createnewdocument;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.SaveFormat;
import com.help.TestFile;
import org.testng.annotations.Test;

/**
 * @Auther: zch
 * @Date: 2019/1/11 17:16
 * @Description:aspose 方式创建word
 */
public class AsposeNewDocument {

    /**
     * aspose 方式创建word
     * @throws Exception
     */
    @Test
    public void test()throws Exception{
        final String dataPath = TestFile.getTestDataDir(this.getClass());
        Document doc = new Document();
        // DocumentBuilder provides members to easily add content to a document.

        DocumentBuilder builder = new DocumentBuilder(doc);
        // Write a new paragraph in the document with some text as "Sample Content..."

        builder.setBold(true);

        builder.writeln("Aspose Sample Content for Word file.");

        // Save the document in DOCX format. The format to save as is inferred from the extension of the file name.
        // Aspose.Words supports saving any document in many more formats.
        doc.save(dataPath + "Aspose_NewDoc_Out.docx", SaveFormat.DOCX);
    }

}
