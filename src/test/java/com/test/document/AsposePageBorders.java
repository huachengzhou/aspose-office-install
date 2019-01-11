package com.test.document;

import com.aspose.words.ConvertUtil;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.PageSetup;
import com.help.Utils;
import org.testng.annotations.Test;

/**
 * @Auther: zch
 * @Date: 2019/1/11 13:55
 * @Description:页表 设置
 */
public class AsposePageBorders {

    /**
     * 页表
     * @throws Exception
     */
    @Test
    public void test()throws Exception{
        String dataPath = Utils.getDataDir(this.getClass());
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        PageSetup pageSetup = builder.getPageSetup();
        pageSetup.setTopMargin(ConvertUtil.inchToPoint(0.5));
        pageSetup.setBottomMargin(ConvertUtil.inchToPoint(0.5));
        pageSetup.setLeftMargin(ConvertUtil.inchToPoint(0.5));
        pageSetup.setRightMargin(ConvertUtil.inchToPoint(0.5));
        pageSetup.setHeaderDistance(ConvertUtil.inchToPoint(0.2));
        pageSetup.setFooterDistance(ConvertUtil.inchToPoint(0.2));

        doc.save(dataPath + "AsposePageBorders.docx");
        System.out.println("Done.");
    }

}
