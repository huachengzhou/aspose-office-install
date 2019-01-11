package com.featurescomparison.workingwithdocuments;

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import com.help.TestFile;
import org.testng.annotations.Test;

import java.io.File;

/**
 * @Auther: zch
 * @Date: 2019/1/11 16:08
 * @Description:
 */
public class AsposeConvertToFormats {

    /**
     * Aspose 转换为Html 或者pdf
     * @throws Exception
     */
    @Test
    public void test()throws Exception{
        final String dataPath = TestFile.getTestDataDir(this.getClass());
        File file  = new File(dataPath);
        if (!file.isDirectory()){
            file.mkdir();
        }
        // Load the document from disk.
        Document doc = new Document(file.getParent() + "\\document.doc");

        doc.save(dataPath + "html/Aspose_DocToHTML_Out.html",SaveFormat.HTML); //Save the document in HTML format.
        doc.save(dataPath + "Aspose_DocToPDF_Out.pdf",SaveFormat.PDF); //Save the document in PDF format.
        doc.save(dataPath + "Aspose_DocToTxt_Out.txt",SaveFormat.TEXT); //Save the document in TXT format.
        doc.save(dataPath + "Aspose_DocToJPG_Out.jpg", SaveFormat.JPEG); //Save the document in JPEG format.

        System.out.println("Aspose - Doc file converted in specified formats");

    }

}
