package fileView;

import data.InfoList;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class XLSXSave {
    Workbook workbook;
    File saveFile;
    InfoList infoList;

    public XLSXSave(String pathMainTable, File saveDir, InfoList infoList) throws IOException, InvalidFormatException {
        this.saveFile = saveDir;
        this.infoList = infoList;
        workbook = new XSSFWorkbook(new FileInputStream(pathMainTable + "\\mainTable.xlsx"));
    }

    public void setAllData(){
        ArrayList<CellStyle> cellStyles = new ArrayList<>();
        for(int i = 0; i < workbook.getSheetAt(0).getRow(1).getPhysicalNumberOfCells();i++) {
            cellStyles.add(workbook.getSheetAt(0).getRow(1).getCell(i).getCellStyle());
            workbook.getSheetAt(0).getRow(1).createCell(i);
        }
        for (int i = 0; i < infoList.mainTable.size(); i++){
            workbook.getSheetAt(0).createRow(i+1).createCell(0).setCellValue(i+1);
            workbook.getSheetAt(0).getRow(i+1).getCell(0).setCellStyle(cellStyles.get(0));
            for (int k = 0; k < infoList.mainTable.get(i).size(); k++){
                workbook.getSheetAt(0).getRow(i+1).createCell(k+1).setCellValue(infoList.mainTable.get(i).get(k));
                workbook.getSheetAt(0).getRow(i+1).getCell(k+1).setCellStyle(cellStyles.get(k+1));
            }
        }
        workbook.getSheetAt(0).setAutoFilter(new CellRangeAddress(0, workbook.getSheetAt(0).getPhysicalNumberOfRows(),
                0, workbook.getSheetAt(0).getRow(0).getPhysicalNumberOfCells() - 1));
    }

    public void close() throws IOException {
        workbook.close();
    }

    public void saveFile() throws IOException {
        workbook.write(new FileOutputStream(saveFile + "\\TRBA466.xlsx"));
    }
}
