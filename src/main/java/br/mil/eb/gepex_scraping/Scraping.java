package br.mil.eb.gepex_scraping;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;

public class Scraping {

	public static void main(String[] args) {
		InputStream is = Scraping.class.getClassLoader().getResourceAsStream("gepex_html_dados.html");
		
		if (is == null) {
            System.out.println("Arquivo não encontrado.");
            return;
        }
		
		try {
			Document doc = Jsoup.parse(is, "UTF-8", "");
			
			// Criar o arquivo Excel
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Dados");
            
    		String[] headers = {
            	"Feito", "Tarefa", "Responsável", "Descrição", "Status", "Início", "Término", 
                "Duração", "Dias", "R$ Planejado", "Total Estimado", "R$ Executado", "R$ Total Executado"
            };
    		
    		Row headerRow = sheet.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
            }
            
            int rowNum = 1;
			
            for (Element element : doc.select(".slick-row")) {
                Row row = sheet.createRow(rowNum++);
                
                // Preencher os dados das colunas de acordo com os seletores
                row.createCell(0).setCellValue(element.select(".l6").text());  // Feito
                row.createCell(1).setCellValue(element.select(".l8").text());  // Tarefa
                row.createCell(2).setCellValue(element.select(".l9").text());  // Responsável
                row.createCell(3).setCellValue(element.select(".l10").text()); // Descrição
                row.createCell(4).setCellValue(element.select(".l11").text()); // Status
                row.createCell(5).setCellValue(element.select(".l12").text()); // Início
                row.createCell(6).setCellValue(element.select(".l13").text()); // Término
                row.createCell(7).setCellValue(element.select(".l14").text()); // Duração
                row.createCell(8).setCellValue(element.select(".l15").text()); // Dias
                row.createCell(9).setCellValue(element.select(".l16").text()); // R$ Planejado
                row.createCell(10).setCellValue(element.select(".l17").text()); // Total Estimado
                row.createCell(11).setCellValue(element.select(".l18").text()); // R$ Executado
                row.createCell(12).setCellValue(element.select(".l19").text()); // R$ Total Executado
            }
            
            try (FileOutputStream fileOut = new FileOutputStream(new File("dados.xlsx"))) {
                workbook.write(fileOut);
            }
            
            workbook.close();
            
            System.out.println("Arquivo Excel gerado com sucesso!");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}