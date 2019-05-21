package org.test;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Table;
import com.google.common.collect.Sets;
import com.help.TestFile;
import com.red.tool.other.MergeCellModel;
import org.testng.annotations.Test;

import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import java.util.UUID;

/**
 * @Auther: zch
 * @Date: 2019/2/14 10:35
 * @Description:
 */
public class Demo4Table {

    @Test
    public void test1()throws Exception{
        List<String> stringList = new ArrayList<String>(12);
        for (int i = 0; i < 12; i++) {
            stringList.add(UUID.randomUUID().toString().substring(0, 4));
        }
        final String dataPath = TestFile.getTestDataParentDir(this.getClass());
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("报酬率测算表");
        Set<MergeCellModel> mergeCellModelList = Sets.newHashSet();
        Table table = builder.startTable();

    }

}
