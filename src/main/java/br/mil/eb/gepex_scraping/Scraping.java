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

	private Workbook workbook;
	private Sheet sheet;

	private final static String[] HEADERS = { "Feito", "Tarefa", "Etapa", "Responsável", "Descrição", "Status",
			"Início", "Término", "Duração", "Dias", "R$ Planejado", "Total Estimado", "R$ Executado",
			"R$ Total Executado" };

	public Workbook getWorkbook() {
		return workbook;
	}

	public Sheet getSheet() {
		return sheet;
	}

	public static String[] getHeaders() {
		return HEADERS;
	}

	public void createExcelFile() {
		workbook = new XSSFWorkbook();
		sheet = workbook.createSheet("Dados");
	}

	public void fillExcelFile(Document doc, Scraping scraping) {
		int rowNum = 1;

		for (Element element : doc.select(".slick-row")) {
			Row row = scraping.getSheet().createRow(rowNum++);

			String etapa = "";
			String tarefaCell = element.select(".l8").text();

			if (tarefaCell.contains("Etapa: ")) {
				String[] partes = tarefaCell.split("Etapa: ");
				if (partes.length > 1) {
					etapa = partes[1].trim();
					etapa = etapa.split(" ")[0].trim();
				}
			}

			String[] values = { element.select(".l6").text(), // Feito
					tarefaCell, // Tarefa
					etapa, // Etapa
					element.select(".l9").text(), // Responsável
					element.select(".l10").text(), // Descrição
					element.select(".l11").text(), // Status
					element.select(".l12").text(), // Início
					element.select(".l13").text(), // Término
					element.select(".l14").text(), // Duração
					element.select(".l15").text(), // Dias
					element.select(".l16").text(), // R$ Planejado
					element.select(".l17").text(), // Total Estimado
					element.select(".l18").text(), // R$ Executado
					element.select(".l19").text() // R$ Total Executado
			};

			for (int i = 0; i < values.length; i++) {
				row.createCell(i).setCellValue(values[i]);
			}
		}
	}
	
	public void saveToFile(String filePath) {
	    try (FileOutputStream fileOut = new FileOutputStream(new File(filePath))) {
	        workbook.write(fileOut);
	        workbook.close();
	        System.out.println("Arquivo Excel gerado com sucesso!");
	    } catch (IOException e) {
	        System.err.println("Erro ao salvar o arquivo: " + e.getMessage());
	    }
	}
	
	public void createHeaders(Scraping scraping) {
		Row headerRow = scraping.getSheet().createRow(0);
		for (int i = 0; i < getHeaders().length; i++) {
			Cell cell = headerRow.createCell(i);
			cell.setCellValue(HEADERS[i]);
		}
	}
	
	private static Document loadHtmlFile(String filePath) {
	    try (InputStream is = Scraping.class.getClassLoader().getResourceAsStream(filePath)) {
	        if (is == null) {
	            throw new IOException("Arquivo não encontrado: " + filePath);
	        }
	        return Jsoup.parse(is, "UTF-8", "");
	    } catch (IOException e) {
	        throw new RuntimeException("Erro ao carregar HTML", e);
	    }
	}

	public static void main(String[] args) {
		final Scraping scraping = new Scraping();
		
		Document doc = loadHtmlFile("gepex_05-03-2025.html");
		scraping.createExcelFile();
		scraping.createHeaders(scraping);
		scraping.fillExcelFile(doc, scraping);
		scraping.saveToFile("dados.xlsx");
	}
}