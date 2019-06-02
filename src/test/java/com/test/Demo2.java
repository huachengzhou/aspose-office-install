package com.test;

import com.aspose.words.Document;
import com.help.Utils;
import org.testng.annotations.Test;
import tool.utils.word.aspose.XmlMailMergeDataTable;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.File;

/**
 * xml数据替换
 * @createDate 2019/1/10
 **/
public class Demo2 {

    @Test
    public void test2() {
        //https://www.cnblogs.com/asxinyu/p/3242754.html#_labelTop

        //https://www.cnblogs.com/itljf/p/5859445.html

        //https://blog.csdn.net/tt871911/article/details/79194909

        //https://www.cnblogs.com/lijiasnong/p/5102363.html
    }

    @Test
    public void test3()throws Exception {

        String dataDir = Utils.getDataDir(Demo2.class);
        System.out.println(dataDir);

        String dataPath = dataDir.toString()+"data\\";
        File file = new File(dataPath);
        if (file.isDirectory()){
            file.mkdir();
        }

        // Use DocumentBuilder from the javax.xml.parsers package and Document class from the org.w3c.dom package to read
        // the XML data file and store it in memory.
        DocumentBuilder db = DocumentBuilderFactory.newInstance().newDocumentBuilder();

        // Parse the XML data.(数据来源)
        org.w3c.dom.Document xmlData = db.parse(dataDir + "Customers.xml");

        // Open a template document.(模板,必须存在)
        Document doc = new Document(dataPath + "mergeDoc.doc");

        // 替换
        // Note that this class also works with a single repeatable region (and any nested regions).(请注意，该类还可用于单个可重复区域（以及任何嵌套区域）)
        // To merge multiple regions at the same time from a single XML data source, use the XmlMailMergeDataSet class.(若要同时从单个XML数据源合并多个区域，请使用 XmlMailMergeDataSet)
        // e.g doc.getMailMerge().executeWithRegions(new XmlMailMergeDataSet(xmlData));
        doc.getMailMerge().execute(new XmlMailMergeDataTable(xmlData, "customer"));

        // Save the output document.
        doc.save(dataPath + "AsposeMailMerge.doc");

        System.out.println("Process Completed Successfully.");
    }

}
