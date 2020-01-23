package com.mmelo.excel.processor;

import com.mmelo.excel.config.ColumnsConfig;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.*;

@Component
@RequiredArgsConstructor
public class GenerateExcelWithImageCell {

    private static final String fileName = "/home/mauriciomelo/excel/novo69.xls";
    private static final String imagePath = "/home/mauriciomelo/excel/rede-logo.png";

    private final ColumnsConfig columnsConfig;

    public void create() throws Exception {
        Workbook wb = new XSSFWorkbook();

        CellStyle styleVertAlingTop = wb.createCellStyle();
        styleVertAlingTop.setVerticalAlignment(VerticalAlignment.BOTTOM);

        Sheet sheet = wb.createSheet();
        sheet.setColumnWidth(0, 30 * 256); //30 default characters width

        Row row = sheet.createRow(0);
        row.setHeight((short)(100 * 20)); //100pt height * 20 = twips (twentieth of an inch point)

        Cell cell = row.createCell(0);
        cell.setCellValue("Replace Kemplon-Pipe");
        cell.setCellStyle(styleVertAlingTop);

        InputStream is = new FileInputStream(imagePath);
        byte[] bytes = IOUtils.toByteArray(is);
        int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
        is.close();

        int left = 20; // 20px
        int top = 20; // 20pt
        int width = Math.round(sheet.getColumnWidthInPixels(0) - left - left); //width in px
        int height = Math.round(row.getHeightInPoints() - top - 10/*pt*/); //height in pt

        drawImageOnExcelSheet((XSSFSheet)sheet, 0, 0, left, top, width, height, pictureIdx);

        FileOutputStream out =
                    new FileOutputStream(new File(GenerateExcelWithImageCell.fileName));
            wb.write(out);
            out.close();
            System.out.println("Arquivo Excel criado com sucesso!");
    }

    private void drawImageOnExcelSheet(XSSFSheet sheet, int row, int col,
                                              int left/*in px*/, int top/*in pt*/, int width/*in px*/, int height/*in pt*/, int pictureIdx) throws Exception {

        final CreationHelper helper = sheet.getWorkbook().getCreationHelper();
        final Drawing drawing = sheet.createDrawingPatriarch();

        final ClientAnchor anchor = helper.createClientAnchor();
        anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);

        anchor.setCol1(col); //first anchor determines upper left position
        anchor.setRow1(row);
        anchor.setDx1(Units.pixelToEMU(left)); //dx = left in px
        anchor.setDy1(Units.toEMU(top)); //dy = top in pt

        anchor.setCol2(col); //second anchor determines bottom right position
        anchor.setRow2(row);
        anchor.setDx2(Units.pixelToEMU(left + width)); //dx = left + wanted width in px
        anchor.setDy2(Units.toEMU(top + height)); //dy= top + wanted height in pt

        drawing.createPicture(anchor, pictureIdx);

    }

}
