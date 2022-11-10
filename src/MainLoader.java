import data.InfoList;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.ArrayList;

public class MainLoader extends JFrame {
    Workbook workbook;
    public MainLoader(File file) throws IOException, InvalidFormatException {
        //String filePath = file.getPath();
        //workbook = new XSSFWorkbook(new FileInputStream(filePath));
    }

    public void saveFile(File saveSample) throws IOException {
        workbook.write(new FileOutputStream(saveSample));
    }
}

