package com.mmelo.excel.processor;

import com.mmelo.excel.config.ColumnsConfig;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.*;
import java.time.LocalDate;

@Component
@RequiredArgsConstructor
public class GenerateExcelWithImageCell {

    private static final String fileName = "/home/mauriciomelo/excel/30.xls";
    private static final String imagePath = "/home/mauriciomelo/excel/rede-logo.png";

    public void create(final LocalDate startDate, final LocalDate endDate) throws Exception {
        final Workbook wb = new XSSFWorkbook();

        final CellStyle styleVertAlingBottom = wb.createCellStyle();
        styleVertAlingBottom.setVerticalAlignment(VerticalAlignment.BOTTOM);

        final Sheet sheet = wb.createSheet();
        sheet.setColumnWidth(0, 30 * 256); //30 default characters width

        final Row row = sheet.createRow(0);
        row.setHeight((short)(1200)); //100pt height * 20 = twips (twentieth of an inch point)

        final Cell cell = row.createCell(0);
        cell.setCellValue("PERIODO: " + startDate.toString() + " A " + endDate.toString()
                + "\n" + "DATA DE EMISSÃO: " + LocalDate.now().toString());
        cell.setCellStyle(styleVertAlingBottom);

        final InputStream is = new FileInputStream(imagePath);
        final int pictureIdx = wb.addPicture(IOUtils.toByteArray(is), Workbook.PICTURE_TYPE_JPEG);
        is.close();

        insertImage((XSSFSheet)sheet, pictureIdx);

        final FileOutputStream out = new FileOutputStream(new File(GenerateExcelWithImageCell.fileName));
        wb.write(out);
        out.close();
    }

    private void insertImage(final XSSFSheet sheet, final int pictureIdx) throws Exception {

        final CreationHelper helper = sheet.getWorkbook().getCreationHelper();
        final Drawing drawing = sheet.createDrawingPatriarch();

        final ClientAnchor anchor = helper.createClientAnchor();
        anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);

        anchor.setCol1(0); //first anchor determines upper left position
        anchor.setRow1(0);

        anchor.setCol2(0); //second anchor determines bottom right position
        anchor.setRow2(0);
        anchor.setDx2(Units.pixelToEMU(80)); //dx = left + wanted width in px
        anchor.setDy2(Units.toEMU(22)); //dy= top + wanted height in pt

        drawing.createPicture( anchor, pictureIdx);
    }

}
