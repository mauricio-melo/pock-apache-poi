package com.mmelo.excel.processor;

import com.mmelo.excel.model.User;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

@Component
@RequiredArgsConstructor
public class GenerateExcelWithImageCell {

    private static final String fileName = "/home/mauriciomelo/excel/84.xls";
    private static final String imagePath = "/home/mauriciomelo/excel/rede-logo.png";

    public void create(final LocalDate startDate, final LocalDate endDate) throws Exception {
        //criação da planilha
        final Workbook wb = new XSSFWorkbook();
        final Sheet sheet = wb.createSheet();

        //criação do header com imagem, titulo e datas
        createHeader(wb, sheet, startDate, endDate);

        //inserção dos dados
        insertData(sheet);

        //criação do arquivo
        final FileOutputStream out = new FileOutputStream(new File(GenerateExcelWithImageCell.fileName));
        wb.write(out);
        out.close();
    }

    public void createHeader(final Workbook wb, final Sheet sheet, final LocalDate startDate, final LocalDate endDate) throws Exception {

        //largura das colunas que serão mescladas
        sheet.setColumnWidth(0, 30 * 130); //15 default characters width
        sheet.setColumnWidth(1, 30 * 130); //15 default characters width

        //altura da linha
        final Row row = sheet.createRow(0);
        row.setHeight((short)(1400)); //100pt height * 20 = twips (twentieth of an inch point)

        // criação do texto
        Font fontBold = wb.createFont();
        fontBold.setFontHeight((short) 200);
        fontBold.setFontName("Liberation Sans");
        fontBold.setBold(true);
        RichTextString richString = new XSSFRichTextString("EXTRATO PARA SIMPLES CONFERÊNCIA"
                + "\n" + "PERIODO: " + startDate.toString() + " A " + endDate.toString()
                + "\n" + "DATA DE EMISSÃO: " + LocalDate.now().toString());
        richString.applyFont(0, 32, fontBold);

        //inserção do texto na celular, com alinhamento inferior
        final Cell cell = row.createCell(0);
        cell.setCellValue(richString);
        final CellStyle styleBottom = wb.createCellStyle();
        styleBottom.setVerticalAlignment(VerticalAlignment.BOTTOM);
        cell.setCellStyle(styleBottom);

        //mescla das celulas
        sheet.addMergedRegion(new CellRangeAddress(0,0,0,1));

        //inserção da imagem
        insertImage((XSSFSheet)sheet, wb);
    }

    private void insertImage(final XSSFSheet sheet, final Workbook wb) throws Exception {
        //captura da imagem
        final InputStream is = new FileInputStream(imagePath);
        final int pictureIdx = wb.addPicture(IOUtils.toByteArray(is), Workbook.PICTURE_TYPE_JPEG);
        is.close();

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

        drawing.createPicture(anchor, pictureIdx);
    }

    private void insertData(Sheet sheet) {
        List<User> listUser = new ArrayList<>();
        listUser.add(new User("Mauricio", "Aprovado"));
        listUser.add(new User("Raphael",  "Aprovado"));
        listUser.add(new User("Alexandre",  "Aprovado"));

        int rownum = 1;
        Row row = sheet.createRow(rownum);

        Cell cellSer1 = row.createCell(0);
        String stringCellValueSer1 = "Nome";
        cellSer1.setCellValue(stringCellValueSer1);

        Cell cellnf1 = row.createCell(1);
        String stringCellValue1 = "Status";
        cellnf1.setCellValue(stringCellValue1);

        rownum ++;
        for (User user : listUser) {
            Row row1 = sheet.createRow(rownum++);
            int cellnum = 0;
            Cell cellNome = row1.createCell(cellnum++);
            cellNome.setCellValue(user.getNome());
            Cell cellRa = row1.createCell(cellnum++);
            cellRa.setCellValue(user.getStatus());
        }
    }

}
