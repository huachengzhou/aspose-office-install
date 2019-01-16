package org.test;

import com.aspose.words.*;
import com.help.TestFile;
import org.apache.commons.lang3.StringUtils;
import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.client.HttpClient;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.util.EntityUtils;
import org.jsoup.Jsoup;
import org.testng.annotations.Test;

import javax.imageio.ImageIO;
import java.awt.*;
import java.io.File;
import java.util.UUID;

/**
 * @Description
 * @createDate 2019/1/15
 **/
public class Demo1 {

    @Test
    public void testA()throws Exception{

        final String dataPath = TestFile.getTestDataParentDir(this.getClass());

        HttpClient httpClient = HttpClients.createDefault();

        HttpPost httpPost = new HttpPost("https://www.cnblogs.com/");
        httpPost.setHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.67 Safari/537.36");
        HttpResponse httpResponse = null;

        httpResponse = httpClient.execute(httpPost);
        if (httpResponse == null) {
            return;
        }
        HttpEntity httpEntity = httpResponse.getEntity();
        if (httpEntity == null) {
            return;
        }
        String webHtml = EntityUtils.toString(httpEntity, "utf-8");
        if (StringUtils.isEmpty(webHtml)) {
            return;
        }
        org.jsoup.nodes.Document document = Jsoup.parse(webHtml);
        org.jsoup.select.Elements post_list = document.select("#post_list .post_item .post_item_body > p");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        //获取Font对象，关于文字的大小，颜色，字体等等基本都在这个里面
        com.aspose.words.Font font = builder.getFont();
        //字体大小
        font.setSize(22);
        //是否粗体
        font.setBold(false);
        //下划线样式，None为无下划线
        font.setUnderline(Underline.DASH);
        //字体颜色
        font.setColor(Color.black);
        //设置字体
        font.setNameFarEast("宋体");
        for (org.jsoup.nodes.Element ele:post_list){
            builder.writeln(ele.text());
        }

        //插入条形码
        int numPages = 4;
        insertBarcodeIntoFooter(builder, doc.getFirstSection(), 1, HeaderFooterType.FOOTER_PRIMARY,dataPath);
        int i = 1;
        while (i < numPages)
        {
            Section cloneSection = (Section) doc.getFirstSection().deepClone(false);
            cloneSection.getPageSetup().setSectionStart(SectionStart.NEW_PAGE);
            doc.appendChild(cloneSection);
            insertBarcodeIntoFooter(builder, cloneSection, i, HeaderFooterType.FOOTER_PRIMARY,dataPath);
            i += 1;
        }

        doc.save(String.format("%s%s%s",dataPath, UUID.randomUUID().toString().substring(1,9),".docx"));
    }

    private void insertBarcodeIntoFooter(DocumentBuilder builder, Section section,
                                         int pageId, int footerType,String dataPath) throws Exception {
        // Move to the footer type in the specific section.
        builder.moveToSection(section.getDocument().indexOf(section));
        builder.moveToHeaderFooter(footerType);

        // Insert the barcode, then move to the next line and insert the ID
        // along with the page number.
        // Use pageId if you need to insert a different barcode on each page. 0
        // = First page, 1 = Second page etc.
        builder.insertImage(ImageIO.read(new File(dataPath + "barcode.png")));
        builder.writeln();
        builder.write("1234567890");
        builder.insertField("PAGE");

        // Create a right aligned tab at the right margin.
        double tabPos = section.getPageSetup().getPageWidth()
                - section.getPageSetup().getRightMargin() - section.getPageSetup().getLeftMargin();
        builder.getCurrentParagraph().getParagraphFormat().getTabStops()
                .add(new TabStop(tabPos, TabAlignment.RIGHT, TabLeader.NONE));

        // Move to the right hand side of the page and insert the page and page
        // total.
        builder.write(ControlChar.TAB);
        builder.insertField("PAGE");
        builder.write(" of ");
        builder.insertField("NUMPAGES");
    }

}
