package com.crossword.powerpoint;

import java.awt.Color;
import java.awt.Rectangle;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.util.Collections;

import org.apache.poi.hslf.usermodel.HSLFAutoShape;
import org.apache.poi.hslf.usermodel.HSLFPictureData;
import org.apache.poi.hslf.usermodel.HSLFPictureShape;
import org.apache.poi.hslf.usermodel.HSLFShape;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFTable;
import org.apache.poi.hslf.usermodel.HSLFTableCell;
import org.apache.poi.hslf.usermodel.HSLFTextBox;
import org.apache.poi.hslf.usermodel.HSLFTextParagraph;
import org.apache.poi.hslf.usermodel.HSLFTextRun;

import org.apache.poi.sl.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Font;

import com.crossword.driver.Preferences;
import com.crossword.gameobjects.CrosswordBoard;
import com.crossword.gameobjects.Letter;

import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.sl.usermodel.TableCell.BorderEdge;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;

public class Powerpoint {
	private HSLFSlideShow ppt;

	public Powerpoint() {
		ppt = new HSLFSlideShow();

	}

	public HSLFSlide createPuzzleSlide(int count) {
		int shapeWidth = 500;
		int shapeHeight = 100;
		int x = (int) ((this.ppt.getPageSize().getWidth() / 2) - (shapeWidth / 2));
		int y = 10;

		HSLFSlide slide = this.ppt.createSlide();
		HSLFTextBox title = slide.createTextBox();
		HSLFTextParagraph p = title.getTextParagraphs().get(0);
		HSLFTextRun r = p.getTextRuns().get(0);
		r.setFontColor(Preferences.TITLE_FONT_COLOR);
		r.setText(Preferences.TITLE_NAME);
		r.setBold(Preferences.TITLE_BOLD);
		r.setFontSize(Preferences.TITLE_FONT_SIZE);
		title.setAnchor(new Rectangle(x, y, shapeWidth, shapeHeight));

		HSLFPictureData logo = null;
		try {
			logo = this.ppt.addPicture(new File(Preferences.directory + "/docs/Picture1.png"),
					PictureData.PictureType.PNG);
		} catch (IOException e) {
			e.printStackTrace();
		}
		HSLFPictureShape shape = new HSLFPictureShape(logo);
		shape.setAnchor(new Rectangle(0, 0, Preferences.LOGO_WIDTH, Preferences.LOGO_HEIGHT));
		slide.addShape(shape);

		HSLFTextBox slideNumber = slide.createTextBox();
		slideNumber.setText("" + count);
		slideNumber.setHorizontalCentered(true);
		slideNumber.setVerticalAlignment(VerticalAlignment.MIDDLE);
		slideNumber.setAnchor(new Rectangle(310, 0, 100, 100));
		HSLFTextParagraph tp = slideNumber.getTextParagraphs().get(0);
		tp.setTextAlign(TextAlign.CENTER);
		HSLFTextRun tr = tp.getTextRuns().get(0);
		tr.setFontSize(Preferences.SLIDE_NUMBER_FONTSIZE);
		tr.setBold(Preferences.SLIDE_NUMBER_FONTBOLD);
		tr.setFontColor(Preferences.SLIDE_NUMBER_COLOR);
		return slide;
	}

	public HSLFSlide createClueSlide(int count) {
		int shapeWidth = 500;
		int shapeHeight = 100;
		int x = (int) ((this.ppt.getPageSize().getWidth() / 2) - (shapeWidth / 2));
		int y = 10;

		HSLFSlide slide = this.ppt.createSlide();
		HSLFTextBox title = slide.createTextBox();
		HSLFTextParagraph p = title.getTextParagraphs().get(0);
		HSLFTextRun r = p.getTextRuns().get(0);
		r.setFontColor(Preferences.TITLE_FONT_COLOR);
		r.setText(Preferences.TITLE_CLUE_NAME);
		r.setBold(Preferences.TITLE_BOLD);
		r.setFontSize(Preferences.TITLE_FONT_SIZE);
		title.setAnchor(new Rectangle(x, y, shapeWidth, shapeHeight));

		HSLFPictureData logo = null;
		try {
			logo = this.ppt.addPicture(new File(Preferences.directory + "/docs/Picture1.png"),
					PictureData.PictureType.PNG);
		} catch (IOException e) {
			e.printStackTrace();
		}
		HSLFPictureShape shape = new HSLFPictureShape(logo);
		shape.setAnchor(new Rectangle(0, 0, Preferences.LOGO_WIDTH, Preferences.LOGO_HEIGHT));
		slide.addShape(shape);

		HSLFTextBox slideNumber = slide.createTextBox();
		slideNumber.setText("" + count);
		slideNumber.setHorizontalCentered(true);
		slideNumber.setVerticalAlignment(VerticalAlignment.MIDDLE);
		slideNumber.setAnchor(new Rectangle(310, 0, 100, 100));
		HSLFTextParagraph tp = slideNumber.getTextParagraphs().get(0);
		tp.setTextAlign(TextAlign.CENTER);
		HSLFTextRun tr = tp.getTextRuns().get(0);
		tr.setFontSize(Preferences.SLIDE_NUMBER_FONTSIZE);
		tr.setBold(Preferences.SLIDE_NUMBER_FONTBOLD);
		tr.setFontColor(Preferences.SLIDE_NUMBER_COLOR);
		return slide;

	}

	public void createClues(HSLFSlide slide, CrosswordBoard board) {
		int shapeWidth = Preferences.CLUE_BOX_WIDTH;
		int shapeHeight = Preferences.CLUE_BOX_HEIGHT;
		int x = (int) ((this.ppt.getPageSize().getWidth() / 2) - (shapeWidth / 2));
		int y = 200;

		HSLFTextBox textbox = slide.createTextBox();
		textbox.setAnchor(new Rectangle(x, y, shapeWidth, shapeHeight));
		
		HSLFTextParagraph p = textbox.getTextParagraphs().get(0);
		p.setTextAlign(TextAlign.LEFT);
		for (int i = 0; i < board.getWordsOnBoard().size(); i++) {
			p.addTextRun(new HSLFTextRun(p));
			HSLFTextRun tr = p.getTextRuns().get(i);
			tr.setText(board.getWordsOnBoard().get(i).getClueNumber() + ". "
					+ board.getWordsOnBoard().get(i).getClue() + "\n");
			
//			textbox.appendText(board.getWordsOnBoard().get(i).getClueNumber() + ". "
//					+ board.getWordsOnBoard().get(i).getClue() + "\n", false);
		}

		
		for (int i = 0; i < p.getTextRuns().size(); i++) {
			HSLFTextRun r = p.getTextRuns().get(i);
			r.setFontColor(Preferences.CLUE_FONT_COLOR);
			r.setFontSize(Preferences.CLUE_FONT_SIZE);
			
		}
	}

	public HSLFSlide createSolutionSlide(int count) throws IOException {
		int shapeWidth = 500;
		int shapeHeight = 100;
		int x = (int) ((this.ppt.getPageSize().getWidth() / 2) - (shapeWidth / 2));
		int y = 10;

		HSLFSlide slide = this.ppt.createSlide();
		HSLFTextBox title = slide.createTextBox();
		HSLFTextParagraph p = title.getTextParagraphs().get(0);
		HSLFTextRun r = p.getTextRuns().get(0);
		r.setFontColor(Preferences.TITLE_FONT_COLOR);
		r.setText(Preferences.TITLE_SOLUTION_NAME);
		r.setBold(Preferences.TITLE_BOLD);
		r.setFontSize(Preferences.TITLE_FONT_SIZE);
		title.setAnchor(new Rectangle(x, y, shapeWidth, shapeHeight));

		HSLFPictureData logo = this.ppt.addPicture(new File(Preferences.directory + "/docs/Picture1.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape shape = new HSLFPictureShape(logo);
		shape.setAnchor(new Rectangle(0, 0, Preferences.LOGO_WIDTH, Preferences.LOGO_HEIGHT));
		slide.addShape(shape);

		HSLFTextBox slideNumber = slide.createTextBox();
		slideNumber.setText("" + count);
		slideNumber.setHorizontalCentered(true);
		slideNumber.setVerticalAlignment(VerticalAlignment.MIDDLE);
		slideNumber.setAnchor(new Rectangle(310, 0, 100, 100));
		HSLFTextParagraph tp = slideNumber.getTextParagraphs().get(0);
		tp.setTextAlign(TextAlign.CENTER);
		HSLFTextRun tr = tp.getTextRuns().get(0);
		tr.setFontSize(Preferences.SLIDE_NUMBER_FONTSIZE);
		tr.setBold(Preferences.SLIDE_NUMBER_FONTBOLD);
		tr.setFontColor(Preferences.SLIDE_NUMBER_COLOR);
		return slide;
	}

	public HSLFTable createSolutionTable(HSLFSlide slide, int row, int col, Letter[][] board, CrosswordBoard cw) {

		int shapeWidth = Preferences.GRID_WIDTH;
		int shapeHeight = Preferences.GRID_HEIGHT;
		int x = (int) ((this.ppt.getPageSize().getWidth() / 2) - ((Preferences.TABLE_COL_WIDTH * Preferences.TABLE_COL) / 2));
		int y = 110;
		HSLFTable table = slide.createTable(row, col);
		table.setAnchor(new Rectangle(x, y, shapeWidth, shapeHeight));

		for (int i = 0; i < table.getNumberOfRows(); i++) {
			for (int k = 0; k < table.getNumberOfColumns(); k++) {

				HSLFTableCell cell = table.getCell(i, k);
				cell.setVerticalAlignment(VerticalAlignment.MIDDLE);
				cell.setHorizontalCentered(true);
				cell.setBorderColor(BorderEdge.top, Preferences.BOARD_BORDER_NOTINPLAY_COLOR);
				cell.setBorderColor(BorderEdge.right, Preferences.BOARD_BORDER_NOTINPLAY_COLOR);
				cell.setBorderColor(BorderEdge.bottom, Preferences.BOARD_BORDER_NOTINPLAY_COLOR);
				cell.setBorderColor(BorderEdge.left, Preferences.BOARD_BORDER_NOTINPLAY_COLOR);
				
				table.setColumnWidth(k, Preferences.TABLE_COL_WIDTH);
				table.setRowHeight(i, Preferences.TABLE_ROW_HEIGHT);

			}
		}

		for (int i = 0; i < table.getNumberOfRows(); i++) {
			for (int k = 0; k < table.getNumberOfColumns(); k++) {

				HSLFTableCell cell = table.getCell(i, k);
				cell.setVerticalAlignment(VerticalAlignment.MIDDLE);
				cell.setHorizontalCentered(true);

				cell.setText(board[i][k].getLetter().toUpperCase());

				HSLFTextRun textBox = cell.getTextParagraphs().get(0).getTextRuns().get(0);
				textBox.setFontColor(Preferences.BOARD_FONT_COLOR);
				textBox.setFontSize(Preferences.TABLE_FONT_SIZE);
				textBox.setBold(Preferences.BOARD_FONT_BOLD);

				if (!cell.getText().equalsIgnoreCase(" ")) {
					cell.setBorderColor(BorderEdge.top, Preferences.BOARD_BORDER_INPLAY_COLOR);
					cell.setBorderColor(BorderEdge.right, Preferences.BOARD_BORDER_INPLAY_COLOR);
					cell.setBorderColor(BorderEdge.bottom, Preferences.BOARD_BORDER_INPLAY_COLOR);
					cell.setBorderColor(BorderEdge.left, Preferences.BOARD_BORDER_INPLAY_COLOR);
				} else {
					cell.setFillColor(Preferences.BOARD_BACKGROUND_COLOR);
				}
			}
		}
		return table;
	}

	public HSLFTable createPuzzleTable(HSLFSlide slide, int row, int col, Letter[][] board, CrosswordBoard cw) {
		
		int shapeWidth = Preferences.GRID_WIDTH;
		int shapeHeight = Preferences.GRID_HEIGHT - 100;
		int x = (int) ((this.ppt.getPageSize().getWidth() / 2) - ((Preferences.TABLE_COL_WIDTH * Preferences.TABLE_COL) / 2));
		int y = 110;
		HSLFTable table = slide.createTable(row, col);
		table.setAnchor(new Rectangle(x, y, shapeWidth, shapeHeight));

		for (int i = 0; i < table.getNumberOfRows(); i++) {
			for (int k = 0; k < table.getNumberOfColumns(); k++) {
				HSLFTableCell cell = table.getCell(i, k);
				cell.setBorderColor(BorderEdge.top, Preferences.BOARD_BORDER_NOTINPLAY_COLOR);
				cell.setBorderColor(BorderEdge.right, Preferences.BOARD_BORDER_NOTINPLAY_COLOR);
				cell.setBorderColor(BorderEdge.bottom, Preferences.BOARD_BORDER_NOTINPLAY_COLOR);
				cell.setBorderColor(BorderEdge.left, Preferences.BOARD_BORDER_NOTINPLAY_COLOR);

				table.setColumnWidth(k, Preferences.TABLE_COL_WIDTH);
				table.setRowHeight(i, Preferences.TABLE_ROW_HEIGHT);
			}
		}
		
		
		for(int j = 0; j < cw.getWordsOnBoard().size(); j++) {
			table.getCell(cw.getWordsOnBoard().get(j).getRow(), cw.getWordsOnBoard().get(j).getCol()).appendText(cw.getWordsOnBoard().get(j).getClueNumber()+". ", false);
			

		}

		for (int i = 0; i < table.getNumberOfRows(); i++) {
			for (int k = 0; k < table.getNumberOfColumns(); k++) {

				HSLFTableCell cell = table.getCell(i, k);
				cell.setVerticalAlignment(VerticalAlignment.MIDDLE);
				cell.setHorizontalCentered(true);


//				if (i == board[i][k].getWord().getRow() && k == board[i][k].getWord().getCol()) {
//					HSLFTextParagraph p = cell.getTextParagraphs().get(0);
//					
//					HSLFTextRun tr = p.getTextRuns().get(0);
////					tr.setText(board[i][k].getWord().getClueNumber() + ".");
//					
//					tr.setFontSize(20.0);
////					tr.setSuperscript(40);
//					
//					
//				}
				


//				if (!board[i][k].getLetter().equalsIgnoreCase(" ")) {
//					cell.setText(" ");
//				} else {
//					cell.setText(board[i][k].getLetter());
//				}

				if (!board[i][k].getLetter().equalsIgnoreCase(" ")) {
					cell.setBorderColor(BorderEdge.top, Preferences.BOARD_BORDER_INPLAY_COLOR);
					cell.setBorderColor(BorderEdge.right, Preferences.BOARD_BORDER_INPLAY_COLOR);
					cell.setBorderColor(BorderEdge.bottom, Preferences.BOARD_BORDER_INPLAY_COLOR);
					cell.setBorderColor(BorderEdge.left, Preferences.BOARD_BORDER_INPLAY_COLOR);
				} else {
					cell.setFillColor(Preferences.BOARD_BACKGROUND_COLOR);
				}

			}
		}
		return table;
	}

	public void writeToPpt(HSLFSlideShow ppt, String fileName) {
		try {
			FileOutputStream out = new FileOutputStream(fileName);
			ppt.write(out);
			out.close();
			ppt.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	public HSLFSlideShow getPpt() {
		return ppt;
	}

	public void setPpt(HSLFSlideShow ppt) {
		this.ppt = ppt;
	}

}