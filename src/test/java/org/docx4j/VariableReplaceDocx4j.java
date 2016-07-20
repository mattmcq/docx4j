package org.docx4j;


import java.io.FileReader;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import com.google.common.reflect.TypeToken;
import com.google.gson.Gson;
import com.google.gson.stream.JsonReader;
import org.docx4j.XmlUtils;
import org.docx4j.convert.out.flatOpcXml.FlatOpcXmlCreator;
import org.docx4j.jaxb.Context;
import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.model.structure.HeaderFooterPolicy;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.contenttype.ContentTypeManager;
import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.JaxbXmlPart;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Document;
import org.docx4j.wml.Ftr;
import org.docx4j.wml.Hdr;
import org.docx4j.wml.Styles;

import javax.xml.bind.JAXBContext;
import javax.xml.bind.Marshaller;
import javax.xml.bind.Unmarshaller;


/**
 * There are at least 3 approaches for replacing variables in
 * a docx.
 *
 * 1. as shows in this example
 * 2. using Merge Fields (see org.docx4j.model.fields.merge.MailMerger)
 * 3. binding content controls to an XML Part (via XPath)
 *
 * Approach 3 is the recommended one when using docx4j. See the
 * ContentControl* examples, Getting Started, and the subforum.
 *
 * Approach 1, as shown in this example, works in simple cases
 * only.  It won't work if your KEY is split across separate
 * runs in your docx (which often happens), or if you want
 * to insert images, or multiple rows in a table.
 *
 * You're encouraged to investigate binding content controls
 * to an XML part.  There is org.docx4j.model.datastorage.migration.FromVariableReplacement
 * to automatically convert your templates to this better
 * approach.
 *
 * OK, enough preaching.  If you want to use VariableReplace,
 * your variables should be appear like so: ${key1}, ${key2}
 *
 * And if you are having problems with your runs being split,
 * VariablePrepare can clean them up.
 *
 */
public class VariableReplaceDocx4j {

    public static void main(String[] args) throws Exception {

        // Exclude context init from timing
        org.docx4j.wml.ObjectFactory foo = Context.getWmlObjectFactory();

        JsonReader jsonReader = new JsonReader(new FileReader("data.json"));


        // Input docx has variables in it: ${colour}, ${icecream}
//        String jinputfilepath = System.getProperty("user.dir") + "/sample-docs/word/unmarshallFromTemplateExample.docx";
//        String inputfilepath = "ContractTemplate-docx4j.docx";
        String inputfilepath = "/Users/mjmc/projects/docx4j/spectrum-template.docx";

        boolean save = true;
        String outputfilepath = System.getProperty("user.dir")
                + "/OUT_VariableReplace.docx";
        System.out.println("outputfilepath = " + outputfilepath);
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage
                .load(new java.io.File(inputfilepath));


        // per stackoverflow  http://stackoverflow.com/questions/17093781/docx4j-does-not-replace-variables/17143488
        VariablePrepare.prepare(wordMLPackage);

        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
//        HeaderPart headerPart = wordMLPackage.getHeaderFooterPolicy().getDefaultHeader();

//        HashMap<String, String> mappings = new HashMap<String, String>();
//        mappings.put("colour", "green");
//        mappings.put("icecream", "chocolate");
        HashMap<String,String> mappings = new Gson().fromJson(jsonReader, new TypeToken<HashMap<String, String>>(){}.getType());
        System.out.println("mappings = " + mappings);

        long start = System.currentTimeMillis();

        // Approach 1 (from 3.0.0; faster if you haven't yet caused unmarshalling to occur):


        List<SectionWrapper> sectionWrappers = wordMLPackage.getDocumentModel().getSections();

        HeaderPart headerPart;
        FooterPart footerPart;
        for (SectionWrapper sw : sectionWrappers) {
            HeaderFooterPolicy hfp = sw.getHeaderFooterPolicy();

            if (hfp.getFirstHeader()!=null) {
                headerPart = hfp.getFirstHeader();
                headerPart.variableReplace(mappings, "{", "}");
            }
            if (hfp.getDefaultHeader()!=null){
                headerPart = hfp.getDefaultHeader();
                headerPart.variableReplace(mappings, "{", "}");
            }
            if (hfp.getEvenHeader()!=null){
                headerPart = hfp.getEvenHeader();
                headerPart.variableReplace(mappings, "{", "}");
            }

            if (hfp.getFirstFooter()!=null) {
                footerPart = hfp.getFirstFooter();
                footerPart.variableReplace(mappings, "{", "}");
            }
            if (hfp.getDefaultFooter()!=null){
                footerPart = hfp.getDefaultFooter();
                footerPart.variableReplace(mappings, "{", "}");
            }
            if (hfp.getEvenHeader()!=null){
                footerPart = hfp.getEvenFooter();
                footerPart.variableReplace(mappings, "{", "}");
            }
        }

        documentPart.variableReplace(mappings, "{", "}");

/*		// Approach 2 (original)

			// unmarshallFromTemplate requires string input
			String xml = XmlUtils.marshaltoString(documentPart.getJaxbElement(), true);
			// Do it...
			Object obj = XmlUtils.unmarshallFromTemplate(xml, mappings);
			// Inject result into docx
			documentPart.setJaxbElement((Document) obj);
*/

        long end = System.currentTimeMillis();
        long total = end - start;
        System.out.println("Time: " + total);

        // Save it
        if (save) {
            SaveToZipFile saver = new SaveToZipFile(wordMLPackage);
            saver.save(outputfilepath);
        } else {
            System.out.println(XmlUtils.marshaltoString(documentPart.getJaxbElement(), true,
                    true));
        }
    }

}