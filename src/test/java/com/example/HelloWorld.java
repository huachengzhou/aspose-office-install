package com.example;

import com.aspose.words.Document;
import com.aspose.words.ReportingEngine;
import com.help.Utils;
import com.linq.Sender;
import org.testng.annotations.Test;

/**
 * @createDate 2019/1/10
 **/
public class HelloWorld {

    @Test
    public void testA()throws Exception{
        //ExStart:HelloWorld
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(HelloWorld.class);
        dataDir = "E:\\doc\\word\\" ;

        String fileName = "HelloWorld.doc";
        // Load the template document.
        Document doc = new Document();

        // Create an instance of sender class to set it's properties.
        Sender sender = new Sender();
        sender.setName("LINQ Reporting Engine");
        sender.setMessage("Hello World");

        // Create a Reporting Engine.
        ReportingEngine engine = new ReportingEngine();

        // Execute the build report.
        engine.buildReport(doc, sender, "sender");


        // Save the finished document to disk.
        doc.save(dataDir + fileName);
        //ExEnd:HelloWorld

        System.out.println("\nTemplate document is populated with the data about the sender.\nFile saved at " + dataDir);
    }

}
