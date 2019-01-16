package org.test;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.help.TestFile;
import org.testng.annotations.Test;

import java.util.UUID;

/**
 * @Author noatn
 * @Description
 * @createDate 2019/1/16
 **/
public class Demo2Table {

    @Test
    public void testE()throws Exception{
        final String dataPath = TestFile.getTestDataParentDir(this.getClass());

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startTable();

        //自动计算行数(第一行)
        builder.insertCell();
        builder.writeln(String.format("%s%s%s",UUID.randomUUID().toString(),",", UUID.randomUUID().toString()));
        builder.insertCell();
        builder.writeln(String.format("%s%s%s",UUID.randomUUID().toString(),",", UUID.randomUUID().toString()));
        builder.endRow();

        //第二行
        builder.insertCell();
        builder.writeln(String.format("%s%s%s",UUID.randomUUID().toString(),",", UUID.randomUUID().toString()));
        builder.insertCell();
        builder.writeln(String.format("%s%s%s",UUID.randomUUID().toString(),",", UUID.randomUUID().toString()));
        builder.endRow();

        //第三行
        builder.insertCell();
        builder.writeln(String.format("%s%s%s",UUID.randomUUID().toString(),",", UUID.randomUUID().toString()));
        builder.insertCell();
        //这设置下单元格样式
        builder.getCellFormat().getBorders().getLeft().setLineWidth(4.0);
        builder.getCellFormat().getBorders().getRight().setLineWidth(4.0);
        builder.getCellFormat().getBorders().getTop().setLineWidth(4.0);
        builder.getCellFormat().getBorders().getBottom().setLineWidth(4.0);
        builder.writeln(String.format("%s%s%s",UUID.randomUUID().toString(),",", UUID.randomUUID().toString()));
        builder.endRow();

        builder.endTable();

        doc.save(String.format("%s%s%s%s",dataPath,this.getClass().getSimpleName(), UUID.randomUUID().toString().substring(1,12),".docx"));
    }

}
