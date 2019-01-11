package com.test.text;

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import com.help.Utils;
import org.testng.annotations.Test;

import java.util.regex.Pattern;

/**
 * @Auther: zch
 * @Date: 2019/1/11 15:15
 * @Description:替换字符串
 */
public class AsposeFindnReplace {

    /**
     * 替换字符串
     * @throws Exception
     */
    @Test
    public void test() throws Exception{
        final String dataPath = Utils.getDataDir(this.getClass());
        Document doc = new Document(dataPath + "replaceDoc.doc");

        // Replaces all 'sad' and 'mad' occurrences with 'bad'
        doc.getRange().replace("sad", "bad", false, true);

        // Replaces all 'sad' and 'mad' occurrences with 'bad'
        doc.getRange().replace(Pattern.compile("[s|m]ad"), "bad");

        doc.save(dataPath + "AsposeReplaced.doc", SaveFormat.DOC);

        System.out.println("Done.");
    }
}
