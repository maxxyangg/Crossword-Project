package com.crossword.driver;

import java.awt.Desktop;
import java.awt.Dimension;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Iterator;
import java.util.Scanner;

import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.json.JSONArray;
import org.json.JSONObject;

import com.crossword.api.APIConnector;
import com.crossword.gameobjects.CrosswordBoard;
import com.crossword.gameobjects.Letter;
import com.crossword.gameobjects.Word;
import com.crossword.powerpoint.Powerpoint;




public class Driver {

	/**
	 * 
	 */

	public static void main(String[] args) {

		/**
		 * wordList will contain the words that the reader object has read. logicalChars
		 * holds the Letter objects of each word
		 */
		ArrayList<String> stringList = new ArrayList<String>();
		ArrayList<Letter> logicalChars = new ArrayList<Letter>();

		/**
		 * wordReader object reads the words contained in the word_list and adds it to
		 * the arrayList wordList
		 */
		Reader reader = new Reader(Preferences.WORD_LIST_FOLDER_LOCATION + Preferences.WORD_LIST_FILE_NAME);
		reader.readFileIntoArray(stringList);

		/**
		 * board represents the skeleton board that will be played on.
		 */
		CrosswordBoard board = new CrosswordBoard();

		board.findBoardSize();
		board.fillBoardWithSpaces();

		Powerpoint crosswordPpt = new Powerpoint();
		crosswordPpt.getPpt().setPageSize(new Dimension(Preferences.SLIDE_WIDTH, Preferences.SLIDE_HEIGHT));
		
		ArrayList<Letter[][]> solutionBoards = new ArrayList<Letter[][]>();
		ArrayList<CrosswordBoard> boards = new ArrayList<CrosswordBoard>();

		int slideCount = 0;
		int numOfPuzzles = 0;
		int repeat = 0;
		int clueNumber = 1;

		Collections.shuffle(stringList);
		Iterator<String> itr = stringList.iterator();
		while (itr.hasNext() && numOfPuzzles < Preferences.MAX_NUMBER_OF_PUZZLES) {
			String currentString = itr.next();
			String[] stringArray = currentString.split(",");

			APIConnector connector = new APIConnector(stringArray[0]);
			connector.getLogicalChars();
			try {
				int responseCode = connector.getConnection().getResponseCode();
				if (responseCode != 200) {
					throw new RuntimeException("HttpResponseCode: " + responseCode);
				} else {
					InputStream inputStream = connector.getConnection().getInputStream();
					Scanner scanner = new Scanner(inputStream);
					String line = new String();
					while (scanner.hasNextLine()) {
						line = scanner.nextLine();
						line = line.substring(line.indexOf("{"));
					}

					JSONObject jsonObj = new JSONObject(line);
					JSONArray jsonArray = jsonObj.getJSONArray("data");

					logicalChars = new ArrayList<Letter>();

					Word word = new Word(stringArray[0], logicalChars, -1, -1, "none", stringArray[1], -1);
					for (int k = 0; k < jsonArray.length(); k++) {
						Letter letter = new Letter(-1, -1, jsonArray.getString(k), word);
						logicalChars.add(letter);
					}
					if (board.insertToBoard(logicalChars, word, clueNumber)) {
						stringList.removeAll(Collections.singleton(currentString));
						itr = stringList.iterator();
						clueNumber++;
						repeat = 0;
					} else {
						Collections.shuffle(stringList);
						itr = stringList.iterator();
						repeat++;
					}
					
					

					if (repeat == 75) {
						HSLFSlide puzzleSlides = crosswordPpt.createPuzzleSlide(++slideCount);
						HSLFSlide clueSlides = crosswordPpt.createClueSlide(slideCount);
//						HSLFSlide solutionsSlides = crosswordPpt.createSolutionSlide(slideCount);

						crosswordPpt.createPuzzleTable(puzzleSlides, Preferences.TABLE_ROW, Preferences.TABLE_COL,
								board.getBoard(), board);


						crosswordPpt.createClues(clueSlides, board);

//						crosswordPpt.createSolutionTable(solutionsSlides, Preferences.TABLE_ROW, Preferences.TABLE_COL,
//								board.getBoard(), board);
						crosswordPpt.writeToPpt(crosswordPpt.getPpt(),
								Preferences.OUTPUT_FOLDER_LOCATION + Preferences.PUZZLE_FILE_NAME);

						numOfPuzzles++;
						
//						for (int i = 0; i < board.getWordsOnBoard().size(); i++) {
//							System.out.println(board.getWordsOnBoard().get(i).getWord());
//						}
						solutionBoards.add(board.getBoard());
						boards.add(board);
						
						board = new CrosswordBoard();
						board.findBoardSize();
						board.fillBoardWithSpaces();
						
						clueNumber = 1;
						


					}else if(!itr.hasNext() && board.getNumberOfWords() > 2) {
						HSLFSlide puzzleSlides = crosswordPpt.createPuzzleSlide(++slideCount);
						HSLFSlide clueSlides = crosswordPpt.createClueSlide(slideCount);
//						HSLFSlide solutionsSlides = crosswordPpt.createSolutionSlide(slideCount);

						crosswordPpt.createPuzzleTable(puzzleSlides, Preferences.TABLE_ROW, Preferences.TABLE_COL,
								board.getBoard(), board);


						crosswordPpt.createClues(clueSlides, board);

//						crosswordPpt.createSolutionTable(solutionsSlides, Preferences.TABLE_ROW, Preferences.TABLE_COL,
//								board.getBoard(), board);
						crosswordPpt.writeToPpt(crosswordPpt.getPpt(),
								Preferences.OUTPUT_FOLDER_LOCATION + Preferences.PUZZLE_FILE_NAME);

						numOfPuzzles++;

//						for (int i = 0; i < board.getWordsOnBoard().size(); i++) {
//							System.out.println(board.getWordsOnBoard().get(i).getWord());
//						}
						solutionBoards.add(board.getBoard());
						boards.add(board);
						
						board = new CrosswordBoard();
						board.findBoardSize();
						board.fillBoardWithSpaces();
						
						clueNumber = 1;

					}
					
					



				}
				connector.disconnect(connector.getConnection());
			} catch (IOException e) {
				e.printStackTrace();
			}

		}
		
		int solutionSlideCount = 1;
		for (int i = 0; i < solutionBoards.size(); i++) {
			try {
				HSLFSlide solutionSlide = crosswordPpt.createSolutionSlide(solutionSlideCount++);
				crosswordPpt.createSolutionTable(solutionSlide, Preferences.TABLE_ROW, Preferences.TABLE_COL,
						solutionBoards.get(i), boards.get(i));
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		crosswordPpt.writeToPpt(crosswordPpt.getPpt(),
				Preferences.OUTPUT_FOLDER_LOCATION + Preferences.PUZZLE_FILE_NAME);
		

		File puzPPT = new File(Preferences.OUTPUT_FOLDER_LOCATION + Preferences.PUZZLE_FILE_NAME);
		Desktop dt = Desktop.getDesktop();
		if (puzPPT.exists()) {
			try {
				dt.open(puzPPT);
			} catch (IOException e) {
				e.printStackTrace();
			}

		}
		
		try {

			File unusedWords = new File(Preferences.WORD_LIST_FOLDER_LOCATION + Preferences.UNUSED_WORDS_FILENAME);
//			FileWriter fileWriter = new FileWriter(unusedWords);
			OutputStreamWriter fileWriter = new OutputStreamWriter(new FileOutputStream(unusedWords),
					"UTF-8");
			
			fileWriter.write("\ufeff");
			for (int i = 0; i < stringList.size(); i++) {
				fileWriter.write(stringList.get(i) + "\n");
			}
			fileWriter.close();
		
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
