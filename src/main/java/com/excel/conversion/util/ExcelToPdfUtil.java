package com.excel.conversion.util;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.common.usermodel.fonts.FontFamily;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Component;
import org.springframework.util.ObjectUtils;
import org.springframework.web.multipart.MultipartFile;

import com.lowagie.text.Chunk;
import com.lowagie.text.Document;
import com.lowagie.text.Font;
import com.lowagie.text.FontFactory;
import com.lowagie.text.PageSize;
import com.lowagie.text.Phrase;
import com.lowagie.text.Rectangle;
import com.lowagie.text.pdf.PdfPCell;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfWriter;

@Component
public class ExcelToPdfUtil {

	public ResponseEntity<?> writeExcelDataInPdf(MultipartFile file,String pageSize) throws Exception {

		try {

			XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
			Sheet sheet = workbook.getSheetAt(0); // Read the first sheet

			// Create a PDF document
			Document document = new Document(PageSize.A1);
			 ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
			PdfWriter.getInstance(document, byteArrayOutputStream);
			document.open();
			

			PdfPTable table = null;

			// Get merged regions from the sheet
			List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();

			for (int i = 0; i <= sheet.getPhysicalNumberOfRows(); i++) {

				if (sheet.getRow(i) != null) {
					List<String> rowList = this.getRowValue(i, sheet);
					table = new PdfPTable(sheet.getRow(i).getPhysicalNumberOfCells());
					table.setWidthPercentage(100);
//				if (ObjectUtils.isNotEmpty(rowList)) {
					// for (Cell cell: sheet.getRow(i)) {
					for (int j = 0; j < sheet.getRow(i).getPhysicalNumberOfCells(); j++) {
						Cell cell = sheet.getRow(i).getCell(j);
						if (cell != null) {
							PdfPCell pdfCell = new PdfPCell();
							XSSFCellStyle style = (XSSFCellStyle) cell.getCellStyle();

							// Handle cell value
							String cellValue = getCellValue(cell);

							// Set the cell value
//							if (StringUtils.isNotEmpty(cellValue)) {
							Phrase p = new Phrase(18, new Chunk(cellValue, setCellFontStyle(style)));

							pdfCell.setPhrase(p);


							// Set borders
							setCellBorders(style, pdfCell);

							// Set alignment
							setCellAlignment(style, pdfCell);

							// back ground set up
							if (style.getFillForegroundColorColor() != null) {
								// byte[] data = (style.getFillBackgroundXSSFColor()).getRGBWithTint();
								System.out.println(style.getFillForegroundColorColor().getARGBHex());
								if (style.getFillForegroundColorColor().getARGBHex() != null)
									pdfCell.setBackgroundColor(
											hex2Rgb(style.getFillForegroundColorColor().getARGBHex()));

							}
//							

							// Check if the current cell is part of a merged region
							boolean isMerged = false;
							int colspan = 1; // Default colspan
							int rowspan = 1; // Default rowspan
							for (CellRangeAddress mergedRegion : mergedRegions) {
								// Check if the current cell is part of a merged region
								if (mergedRegion.isInRange(i, cell.getColumnIndex())) {
									// For colspan
									if (mergedRegion.getFirstRow() == i
											&& mergedRegion.getFirstColumn() == cell.getColumnIndex()) {
										colspan = mergedRegion.getLastColumn() - mergedRegion.getFirstColumn() + 1;
									}

									// For rowspan
									if (mergedRegion.getFirstColumn() == cell.getColumnIndex()
											&& mergedRegion.getFirstRow() == i) {
										rowspan = mergedRegion.getLastRow() - mergedRegion.getFirstRow() + 1;
									}

									isMerged = true;
									break;
								}
							}

							// If the cell is part of a merged region, set colspan and rowspan
							if (isMerged) {
								pdfCell.setColspan(colspan);
								pdfCell.setRowspan(rowspan);
							}

					
							table.addCell(pdfCell);

						} else {
							PdfPCell pdfCell = new PdfPCell();
							pdfCell.setPhrase(new Phrase(""));
							pdfCell.setFixedHeight(i);
							// pdfCell.setBackgroundColor(new Color(192, 192, 192));
							table.addCell(pdfCell);
						}
					}
					// Add the table to the document after processing the row
					document.add(table);
				}

			}

			document.close();
			 //ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
			workbook.close();
			
			System.out.println("Excel data successfully written to PDF.");
			
			return ResponseEntity.ok()
					.header(HttpHeaders.CONTENT_DISPOSITION,
							"attachment; filename=" + file.getOriginalFilename())
					.header(HttpHeaders.ACCESS_CONTROL_EXPOSE_HEADERS, HttpHeaders.CONTENT_DISPOSITION)
					//.contentType(MediaType.parseMediaType("application/octet-stream"))
					.body(byteArrayOutputStream.toByteArray());
			
		} catch (Exception e) {
			System.out.println("exception in pdf generation");
			throw e ;
		}

	}

	private String getCellValue(Cell cell) {
		String value = "";
		switch (cell.getCellType()) {
		case STRING:
			value = cell.getStringCellValue();
			break;
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				value = cell.getDateCellValue().toString();
			} else {
				value = String.valueOf(cell.getNumericCellValue());
			}
			break;
		case BOOLEAN:
			value = String.valueOf(cell.getBooleanCellValue());
			break;
		case FORMULA:
			value = cell.getCellFormula();
			break;
		default:
			break;
		}
		return value;
	}

	private void setCellBorders(CellStyle style, PdfPCell pdfCell) {
		if (style.getBorderTop() != BorderStyle.NONE) {
			pdfCell.setBorderWidthTop(1);
		} else {
			pdfCell.setBorderWidthTop(Rectangle.NO_BORDER);
		}
		if (style.getBorderBottom() != BorderStyle.NONE) {
			pdfCell.setBorderWidthBottom(1);
		} else {
			pdfCell.setBorderWidthBottom(Rectangle.NO_BORDER);
		}
		if (style.getBorderLeft() != BorderStyle.NONE) {
			pdfCell.setBorderWidthLeft(1);
		} else {
			pdfCell.setBorderWidthLeft(Rectangle.NO_BORDER);
		}
		if (style.getBorderRight() != BorderStyle.NONE) {
			pdfCell.setBorderWidthRight(1);
		} else {
			pdfCell.setBorderWidthRight(Rectangle.NO_BORDER);
		}
	}

	private void setCellAlignment(CellStyle style, PdfPCell pdfCell) {
		if (style.getAlignment() == HorizontalAlignment.CENTER) {
			pdfCell.setHorizontalAlignment(PdfPCell.ALIGN_CENTER);
		} else if (style.getAlignment() == HorizontalAlignment.LEFT) {
			pdfCell.setHorizontalAlignment(PdfPCell.ALIGN_LEFT);
		} else if (style.getAlignment() == HorizontalAlignment.RIGHT) {
			pdfCell.setHorizontalAlignment(PdfPCell.ALIGN_RIGHT);
		}

		if (style.getVerticalAlignment() == VerticalAlignment.TOP) {
			pdfCell.setVerticalAlignment(PdfPCell.ALIGN_TOP);
		} else if (style.getVerticalAlignment() == VerticalAlignment.CENTER) {
			pdfCell.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
		} else if (style.getVerticalAlignment() == VerticalAlignment.BOTTOM) {
			pdfCell.setVerticalAlignment(PdfPCell.ALIGN_BOTTOM);
		}
	}

	private Font setCellFontStyle(XSSFCellStyle style) {
		XSSFFont font = style.getFont();

		XSSFColor color = font.getXSSFColor();

		Font font1 = new Font();

		if (!ObjectUtils.isEmpty(color))
			font1.setColor(this.hex2Rgb(color.getARGBHex()));

		System.out.println("Font name : " + font.getFontName());

		font1.setStyle(setFontStyle(font));

		FontFamily family = FontFamily.valueOf(((XSSFFont) font).getFamily());
		System.out.println("Font family : " + family);

		font1.setFamily(FontFactory.TIMES_BOLDITALIC);

		System.out.println("Font family in int : " + font.getFamily());

		System.out.println("Font FontHeight : " + font.getFontHeight());

		font.setFontHeight(font.getFontHeight());

		return font1;

	}

	public static int setFontStyle(XSSFFont font) {
		// Start with the default font style (i.e., normal)
		int fontStyle = Font.NORMAL;

		// Check if the XSSFFont is bold, italic, underlined, or strike-through
		if (font.getBold()) {
			fontStyle = Font.BOLD;
		}
		if (font.getItalic()) {
			fontStyle = Font.ITALIC;
		}

		if (font.getStrikeout()) {
			fontStyle = Font.STRIKETHRU;
		}
		return fontStyle;
	}

	private void setCellBorders(CellStyle style) {
		if (style != null) {
			if (style.getBorderTop() != BorderStyle.NONE) {
				style.setBorderTop(BorderStyle.THIN);
			} else {
				style.setBorderTop(BorderStyle.NONE);
			}
			if (style.getBorderBottom() != BorderStyle.NONE) {
				style.setBorderBottom(BorderStyle.THIN);
			} else {
				style.setBorderBottom(BorderStyle.NONE);
			}
			if (style.getBorderLeft() != BorderStyle.NONE) {
				style.setBorderLeft(BorderStyle.THIN);
			} else {
				style.setBorderLeft(BorderStyle.NONE);
			}
			if (style.getBorderRight() != BorderStyle.NONE) {
				style.setBorderRight(BorderStyle.THIN);
			} else {
				style.setBorderRight(BorderStyle.NONE);
			}
		}
	}

	public Color hex2Rgb(String hex) {

		hex = hex.replace("#", "");

		int alpha = Integer.parseInt(hex.substring(0, 2), 16);
		int red = Integer.parseInt(hex.substring(2, 4), 16);
		int green = Integer.parseInt(hex.substring(4, 6), 16);
		int blue = Integer.parseInt(hex.substring(6, 8), 16);

		return new Color(red, green, blue, alpha);

	}
	/**
	 * This method will return the list of values in the row
	 * @param index
	 * @param sheet
	 * @return
	 */

	public List<String> getRowValue(int index, Sheet sheet) {
		List<String> list = new ArrayList<>();

		if (!ObjectUtils.isEmpty(sheet.getRow(index))) {
			for (Cell cell : sheet.getRow(index)) {

				switch (cell.getCellType()) {
				case STRING:
					list.add(cell.getStringCellValue());
					CellStyle test = cell.getCellStyle();
					break;
				case NUMERIC:
					list.add(String.valueOf(cell.getNumericCellValue()));

					break;
				case BOOLEAN:
					list.add(String.valueOf(cell.getBooleanCellValue()));

					break;
				case FORMULA:
					list.add(cell.getCellFormula().toString());

					break;
				}
			}
		}

		return list;
	}
}
