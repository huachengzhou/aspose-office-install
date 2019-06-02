package com.test.text;

import com.aspose.words.Document;
import com.aspose.words.FontSettings;
import com.help.Utils;
import org.testng.annotations.Test;

/**
 * @Auther: zch
 * @Date: 2019/1/11 15:28
 * @Description:渲染为其它格式内容如pdf等
 */
public class AsposeSpecifyDefaultFontswhileRendering {
    /**
     *
     *渲染为其它格式内容如pdf等
     * @throws Exception
     */
    @Test
    public void test() throws Exception {
        final String dataPath = Utils.getDataDir(this.getClass());
        Document doc = new Document(dataPath + "document.doc");

        //如果在渲染期间找不到此处定义的默认字体，则使用计算机上最近的字体。
        // If the default font defined here cannot be found during rendering then the closest font on the machine is used instead.
//        FontSettings.setDefaultFontName("Arial Unicode MS");
        FontSettings.getDefaultInstance().setDefaultFontName("Arial Unicode MS");

        //现在，在任何渲染调用期间，使用设置的默认字体来替代任何丢失的字体。
        // Now the set default font is used in place of any missing fonts during any rendering calls.
        doc.save(dataPath + "AsposeSetDefaultFont_Out.pdf");
        doc.save(dataPath + "AsposeSetDefaultFont_Out.xps");

        System.out.println("Process Completed Successfully");
    }
}
