package com.test.document;

import com.aspose.words.Document;
import com.help.Utils;
import org.testng.annotations.Test;

/**
 * @Auther: zch
 * @Date: 2019/1/11 11:57
 * @Description:txt文件内容写入文档
 */
public class AsposeLoadTxtFile {

    /**
     * txt文件内容写入文档
     * @throws Exception
     */
    @Test
    public void test()throws Exception{
        String dataPath = Utils.getDataDir(this.getClass());

        // The encoding of the text file is automatically detected.
        Document doc = new Document(dataPath + "LoadTxt.txt");

        // Save as any Aspose.Words supported format, such as DOCX.
        doc.save(dataPath + "AsposeLoadTxt_Out.docx");

        System.out.println("Process Completed Successfully");
    }

}
