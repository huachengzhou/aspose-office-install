package com.test.xml;

import com.aspose.words.Document;
import com.help.Utils;
import com.red.tool.xml.XmlMailMergeDataTable;
import org.testng.annotations.Test;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.File;

/**
 * @Auther: zch
 * @Date: 2019/1/11 16:01
 * @Description:xml数据替换word模板
 */
public class XMLMailMerge {

    @Test
    public void test()throws Exception{
        String dataDir = Utils.getDataDir(this.getClass());
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
