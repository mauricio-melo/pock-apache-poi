package com.mmelo.excel.processor;

import com.mmelo.excel.config.ColumnsConfig;
import com.mmelo.excel.model.User;
import lombok.RequiredArgsConstructor;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

import static org.apache.poi.ss.usermodel.ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE;

@Component
@RequiredArgsConstructor
public class GenerateExcelWithImageWithoutCell {

    private static final String fileName = "/home/mauriciomelo/excel/novo29.xls";
    private static final String imagePath = "/home/mauriciomelo/excel/rede-logo.png";

    private final ColumnsConfig columnsConfig;

    public void create() throws Exception {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("User");

        List<User> listUser = new ArrayList<>();
        listUser.add(new User("Mauricio", "Aprovado"));
        listUser.add(new User("Raphael",  "Aprovado"));
        listUser.add(new User("Alexandre",  "Reprovado1111111111111111111111"));

        int rownum = 6;

//        columnsConfig.getColumns()
//                .forEach(s -> {
//                    Cell cellSer1 = row.createCell(columnsConfig.getColumns().indexOf(s));
//                    cellSer1.setCellValue(s);
//                });

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

        int quantidadeColunas = sheet.getPhysicalNumberOfRows();
        for(int i = 0; i < quantidadeColunas; i++ ) {
            sheet.autoSizeColumn(i);
        }


        //LOGO-------------------------

        final FileInputStream stream =
                new FileInputStream(imagePath);
        final CreationHelper helper = workbook.getCreationHelper();
        final Drawing drawing = sheet.createDrawingPatriarch();

        final ClientAnchor anchor = helper.createClientAnchor();
        anchor.setAnchorType(DONT_MOVE_AND_RESIZE);

        final int pictureIndex =
                workbook.addPicture(IOUtils.toByteArray(stream), Workbook.PICTURE_TYPE_PNG);

        anchor.setCol1( 0 );
        anchor.setRow1( 0 ); // same row is okay
        anchor.setRow2( 5 );
        anchor.setCol2( 5 );

        final Picture pict = drawing.createPicture( anchor, pictureIndex );
        pict.resize(1.0, 3.0);

        //------------------------------------------------




        try {
            FileOutputStream out =
                    new FileOutputStream(new File(GenerateExcelWithImageWithoutCell.fileName));
            workbook.write(out);
            out.close();
            System.out.println("Arquivo Excel criado com sucesso!");

        } catch (FileNotFoundException e) {
            e.printStackTrace();
            System.out.println("Arquivo não encontrado!");
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Erro na edição do arquivo!");
        }

    }

}
