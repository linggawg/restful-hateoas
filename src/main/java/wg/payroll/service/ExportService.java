package wg.payroll.service;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.springframework.stereotype.Component;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

@Component
public class ExportService {
    public void createCell(XSSFRow row, int columnCount, Object value, CellStyle style) {
        XSSFCell cell = row.createCell(columnCount);
        if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof Long) {
            cell.setCellValue((Long) value);
        } else if (value instanceof Double) {
            cell.setCellValue((Double) value);
        } else if (value instanceof Date) {
            DateFormat dateFormatter = new SimpleDateFormat("dd-MM-yyyy");
            String currentDate = dateFormatter.format(value);
            cell.setCellValue(currentDate);
        } else {
            cell.setCellValue((String) value);
        }
        cell.setCellStyle(style);
    }
}
