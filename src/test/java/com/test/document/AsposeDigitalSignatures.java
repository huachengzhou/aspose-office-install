package com.test.document;

import com.aspose.words.FileFormatInfo;
import com.aspose.words.FileFormatUtil;
import com.help.Utils;
import org.testng.annotations.Test;

import java.io.File;

/**
 * @Auther: zch
 * @Date: 2019/1/11 14:28
 * @Description:数字签名
 */
public class AsposeDigitalSignatures {

    /**
     * 数字签名
     * @throws Exception
     */
    @Test
    public void test()throws Exception{
        String dataPath = Utils.getDataDir(this.getClass());
        // The path to the document which is to be processed.
        String filePath = dataPath + "担保书.docx";

        FileFormatInfo info = FileFormatUtil.detectFileFormat(filePath);
        if (info.hasDigitalSignature())
        {
            System.out.println(java.text.MessageFormat.format(
                    "Document {0} has digital signatures, they will be lost if you open/save this document with Aspose.Words.",
                    new File(filePath).getName()));
        }
        else
        {
            System.out.println("Document has no digital signature.");
        }
        System.out.println("Process Completed Successfully");
    }

}
