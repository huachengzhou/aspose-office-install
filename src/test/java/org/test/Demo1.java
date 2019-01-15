package org.test;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
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
        for (org.jsoup.nodes.Element ele:post_list){
            builder.writeln(ele.text());
        }
        doc.save(String.format("%s%s%s",dataPath, UUID.randomUUID().toString().substring(1,9),".docx"));
    }

}
