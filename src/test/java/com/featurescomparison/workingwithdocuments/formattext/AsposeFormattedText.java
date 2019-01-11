package com.featurescomparison.workingwithdocuments.formattext;

import com.aspose.words.*;
import com.aspose.words.Font;
import com.help.TestFile;
import org.testng.annotations.Test;

import java.awt.*;


/**
 * @Auther: zch
 * @Date: 2019/1/11 17:25
 * @Description:格式化
 */
public class AsposeFormattedText {

    /**
     * 格式化
     * @throws Exception
     */
    @Test
    public void test() throws Exception {
        final String dataPath = TestFile.getTestDataDir(this.getClass());
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set paragraph formatting properties
        ParagraphFormat paragraphFormat = builder.getParagraphFormat();
        paragraphFormat.setAlignment(ParagraphAlignment.CENTER);
        paragraphFormat.setLeftIndent(50);
        paragraphFormat.setRightIndent(50);
        paragraphFormat.setSpaceAfter(25);

        // Output text
        builder.writeln("I'm a very nice formatted paragraph. I'm intended to demonstrate how the left and right indents affect word wrapping.");

        // Set font formatting properties
        Font font = builder.getFont();
        font.setBold(true);
        font.setColor(Color.BLUE);
        font.setItalic(true);
        font.setName("Arial");
        font.setSize(24);
        font.setSpacing(5);
        font.setUnderline(Underline.DOUBLE);

        // Output formatted text
        builder.writeln("I'm a very nice formatted string.");

        doc.save(dataPath + "Aspose_FormattedText_Out.doc");

        System.out.println("Process Completed Successfully");
    }


}
