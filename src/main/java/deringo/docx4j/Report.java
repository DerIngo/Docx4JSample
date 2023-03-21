package deringo.docx4j;

import java.util.ArrayList;
import java.util.List;

public class Report {

    public String ueberschrift = "Sample Report";
    public String text = "Lorem Ipsum";
    public List<String[]> tabelle;
    
    public Report() {
        tabelle = new ArrayList<>();
        tabelle.add(new String[] {"Montag", "Regen", "4°C"});
        tabelle.add(new String[] {"Dienstag", "Wolkig", "6°C"});
        tabelle.add(new String[] {"Mittwoch", "Sonne", "10°C"});
    }
}