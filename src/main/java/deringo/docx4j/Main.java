package deringo.docx4j;

import java.awt.Desktop;
import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;

import jakarta.xml.bind.JAXBElement;

public class Main {

    final static String XPATH_TO_SELECT_TEXT_NODES = "//w:t";
    
    public static void main(String[] args) throws Exception {
        File template = getTemplateFile();
        File output   = getOutputFile();
        Report report = new Report();
        
        output = generateReport(template, output, report);

        Desktop.getDesktop().open(output);
    }
    
    public static File generateReport(File template, File output, Report report) throws Exception {
        WordprocessingMLPackage wordPackage = WordprocessingMLPackage.load(template);
        textErsetzen(wordPackage, report);
        tabellenzeilenErzeugen(wordPackage, report);
        wordPackage.save(output);
        return output;
    }
    
    private static void textErsetzen(WordprocessingMLPackage wordPackage, Report report) throws Exception {
        List<Object> texts = wordPackage.getMainDocumentPart().getJAXBNodesViaXPath(XPATH_TO_SELECT_TEXT_NODES, true);
        for (Object obj : texts) {  
            Text text = (Text) ((JAXBElement<?>) obj).getValue();
            
            String textValue;
            if ("ÜBERSCHRIFT1".equals(text.getValue())) {
                textValue = report.ueberschrift;
            } else if ("TEXT1".equals(text.getValue())) {
                textValue = report.text;
            } else {
                textValue = text.getValue();
            }
            
            text.setValue(textValue);
        } 
    }
    
    private static void tabellenzeilenErzeugen(WordprocessingMLPackage wordPackage, Report report) {
        for (Object o : wordPackage.getMainDocumentPart().getContent()) {
            if (o instanceof JAXBElement) {
                @SuppressWarnings("unchecked")
                Tbl tbl = ((JAXBElement<Tbl>)o).getValue();
                String tblCaption = tbl.getTblPr().getTblCaption() != null ? tbl.getTblPr().getTblCaption().getVal() : null;
                if ("MyTabelle1".equals(tblCaption)) {
                    fillTabelle1(tbl, report);
                } else {
                    System.out.println("tblCaption: " + tblCaption);
                }
            }
        }
    }
    
    private static void fillTabelle1(Tbl tbl, Report report) {
        for (String[] zeile : report.tabelle) {
            addRow(tbl, zeile[0], zeile[1], zeile[2]);
        }
    }
    
    private static void addRow(Tbl tbl, String... cells) {
        ObjectFactory factory = new ObjectFactory();
        
        // setup Font
        String fontName = "Arial";
        RFonts fonts = factory.createRFonts();
        fonts.setAscii(fontName);
        fonts.setCs(fontName);
        fonts.setHAnsi(fontName);
        RPr rpr = factory.createRPr();
        rpr.setRFonts(fonts);
        
        // new Row
        Tr tr = factory.createTr();
        for (String cell : cells) {
            Tc tc = factory.createTc();
            
            P p = factory.createP();
            R r = factory.createR();
            r.setRPr(rpr);
            
            Text text = factory.createText();
            
            text.setValue(cell);
            r.getContent().add(text);
            p.getContent().add(r);
            
            tc.getContent().add(p);
            
            tr.getContent().add(tc);
        }
        tbl.getContent().add(tr);
    }
    
    private static File getTemplateFile() throws Exception {
        File template = new File( Main.class.getClassLoader().getResource("template.docx").toURI() );
        return template;
    }

    private static File getOutputFile() throws Exception {
        String tmpdir = System.getProperty("java.io.tmpdir");
        Path tmpfile = Files.createTempFile(Paths.get(tmpdir), "MyReport" + "-", ".docx");
        return tmpfile.toFile();
    }
}
