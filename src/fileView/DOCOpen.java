package fileView;

import data.InfoList;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

public class DOCOpen {
    XWPFDocument workbook;
    public DOCOpen(File file) throws IOException, InvalidFormatException {
        workbook = new XWPFDocument(new FileInputStream(file));
    }

    public void close() throws IOException {
        workbook.close();
    }

    public void getMainInfo(InfoList infoList) throws IOException {
        boolean checkParenthesis = false;
        for(int i = 171; i < workbook.getParagraphs().size() && !workbook.getParagraphs().get(i).getText().equals("(including IJSEM issue July 2008)"); i++){
            if(workbook.getParagraphs().get(i).getText().contains("www.baua.de/abas")){
                i = i + 3;
            }
            if(!workbook.getParagraphs().get(i).getText().equals("")){
                System.out.println(workbook.getParagraphs().get(i).getText());
                String str[] = {};
                String temp = workbook.getParagraphs().get(i).getText();
                str = temp.split(" ");
                if (str.length > 2 && str[1].equals("") && str[0].equals("")){
                    infoList.mainTable.get(infoList.mainTable.size()-1).set(0, infoList.mainTable.get(infoList.mainTable.size()-1).get(0) + " →");
                    for (int k = 2; k < str.length; k++){
                        infoList.mainTable.get(infoList.mainTable.size()-1).set(0, infoList.mainTable.get(infoList.mainTable.size()-1).get(0) + " " + str[k]);
                    }
                } else if (str.length > 2 && str[1].equals("–") && str[0].equals("")){
                    for (int k = 0; k < str.length; k++){
                        infoList.mainTable.get(infoList.mainTable.size()-1).set(0, infoList.mainTable.get(infoList.mainTable.size()-1).get(0) + " " + str[k]);
                    }
                } else if (str.length > 2 && str[1].charAt(0) == ('(') && str[0].equals("")){
                    for (int k = 0; k < str.length; k++){
                        if (k == 0){
                            infoList.mainTable.get(infoList.mainTable.size()-1).set(0, infoList.mainTable.get(infoList.mainTable.size()-1).get(0) + str[k]);
                        } else if (str[k].equals("1")
                                || str[k].equals("2")
                                || str[k].equals("3")
                                || str[k].equals("4")) {
                            infoList.mainTable.get(infoList.mainTable.size()-1).add(str[k]);
                        } else {
                            infoList.mainTable.get(infoList.mainTable.size()-1).set(0, infoList.mainTable.get(infoList.mainTable.size()-1).get(0) + " " + str[k]);
                        }
                    }
                }
                else if (infoList.remarks.contains(str[0]) || infoList.remarks.contains(str[0].replace(",", ""))){
                    infoList.mainTable.get(infoList.mainTable.size()-1).add("");
                    for (int k = 0; k < str.length; k++){
                        if (k == 0){
                            infoList.mainTable.get(infoList.mainTable.size()-1)
                                    .set(infoList.mainTable.get(infoList.mainTable.size()-1).size()-1, str[k].replace(",", ";"));
                        } else {
                            infoList.mainTable.get(infoList.mainTable.size()-1)
                                    .set(infoList.mainTable.get(infoList.mainTable.size()-1).size()-1, infoList.mainTable
                                            .get(infoList.mainTable.size()-1).get(infoList.mainTable
                                                    .get(infoList.mainTable.size()-1).size()-1) + str[k].replace(",", ";"));
                        }
                    }
                } else if (temp.contains("(") && !temp.contains(")") || checkParenthesis){
                    for (int k = 0; k < str.length; k++){
                        infoList.mainTable.get(infoList.mainTable.size()-1).set(0, infoList.mainTable.get(infoList.mainTable.size()-1).get(0) + " " + str[k]);
                    }
                    if (temp.contains(")")) {
                        checkParenthesis = false;
                    } else {
                        checkParenthesis = true;
                    }
                } else if((temp.length() == 1
                        && (str[0].equals("1")
                        || str[0].equals("2")
                        || str[0].equals("3")
                        || str[0].equals("4")))){
                    infoList.mainTable.get(infoList.mainTable.size()-1).add(str[0]);
                } else if (!temp.equals("Species 1 2 3 4")
                        && !temp.equals("Notes")
                        && !temp.equals("G Classification into risk group deviates from „Liste risikobewerteter Spender- und Empfängerorganismen für gentechnische Arbeiten“.")
                        && !temp.contains("G Classification into risk group deviates from „Liste risikobewerteter Spender- und Empfängerorganismen für gentechnische")
                        && !temp.contains("Arbeiten“.")){
                    infoList.mainTable.add(new ArrayList<>());
                    infoList.mainTable.get(infoList.mainTable.size()-1).add("");
                    for (int k = 0; k < str.length; k++){
                        if(str[k].equals("1")
                                || str[k].equals("2") || str[k].equals("3")
                                || str[k].equals("4") || str[k].equals("3(**)")
                                || str[k].equals("2G") || str[k].equals("1G"))
                        {
                            infoList.mainTable.get(infoList.mainTable.size()-1).add(str[k]);
                        } else if(infoList.remarks.contains(str[k].replace(",", "")))
                        {
                            if( infoList.mainTable.get(infoList.mainTable.size()-1).size() < 3){
                                infoList.mainTable.get(infoList.mainTable.size()-1).add(str[k].replace(",", ";"));
                            } else {
                                infoList.mainTable.get(infoList.mainTable.size()-1)
                                        .set(infoList.mainTable.get(infoList.mainTable.size()-1).size()-1,
                                                (infoList.mainTable.get(infoList.mainTable.size()-1)
                                                        .get(infoList.mainTable.get(infoList.mainTable.size()-1).size()-1) + str[k].replace(",", ";")));
                            }
                        } else if(str[k].equals(""))
                        {
                            infoList.mainTable.get(infoList.mainTable.size()-1).set(0,  infoList.mainTable.get(infoList.mainTable.size()-1).get(0) + " →" );
                        } else {
                            if (infoList.mainTable.get(infoList.mainTable.size()-1).get(0).equals("")){
                                infoList.mainTable.get(infoList.mainTable.size()-1).set(0,  str[k]);
                            } else {
                                infoList.mainTable.get(infoList.mainTable.size()-1).set(0,  infoList.mainTable.get(infoList.mainTable.size()-1).get(0) + " " + str[k]);
                            }
                        }
                    }
                }
            }
        }
        for (int i = 0; i < infoList.mainTable.size(); i++){
            System.out.println(infoList.mainTable.get(i));
        }
    }
}