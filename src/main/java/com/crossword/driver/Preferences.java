package com.crossword.driver;

import java.awt.Color;
import java.sql.Timestamp;

public class Preferences {

	// This variable captures the current time in milliseconds.
	private static long now = System.currentTimeMillis();

	// The Timestamp object is used to return a timestamp.
	private static Timestamp timestamp = new Timestamp(now);

	// The String variable 'directory' returns the path of the running project. This
	// is used to get files inside of the project sub-directories
	// Workspace\ICS499-Skeleton_Project
	public static String directory = System.getProperty("user.dir") + "\\src\\main\\resources\\";

	// MAX_NUMBER_OF_WORDS specifies the number of words that appear on each puzzle
//	public final static int MAX_NUMBER_OF_WORDS = 3;

	// MAX_NUMBER_OF_PUZZLES specifies the number of puzzles to create
	public final static int MAX_NUMBER_OF_PUZZLES = 1;

	// WORD_LIST_FILE_NAME specifies the file name of the word list you want to use.
	public final static String WORD_LIST_FILE_NAME = "telugu_crossword_wordlist.csv";

	// PUZZLE_FILE_NAME specifies the name of the output filename for the puzzle
	public final static String PUZZLE_FILE_NAME = "telugu_skeleton_puzzle_" + timestamp.getNanos() / 10000 + ".ppt";

	// SOLUTION_FILE_NAME specifies the name of the output filename for the solution
	// This variable is not used because puzzle and solution have been combined.
//	public final static String SOLUTION_FILE_NAME = "english_skeleton_solution_" + timestamp.getNanos() / 10000
//			+ ".ppt";

	// OUTPUT_FOLDER_LOCATION specifies the location of the output files
	public final static String OUTPUT_FOLDER_LOCATION = directory + "\\PowerPoints\\";

	// WORD_LIST_FOLDER_LOCATION specfies the location of the word list
	public final static String WORD_LIST_FOLDER_LOCATION = directory + "\\docs\\";
	
	// UNUSED_WORDS_FILENAME is the filename for the list of unused words.
	public final static String UNUSED_WORDS_FILENAME = "unused_words.csv";

	// BOARD_FONT_COLOR specifies the font color of the skeleton board
	public final static Color BOARD_FONT_COLOR = Color.black;

	// TABLE_FONT_BOLD toggles the font bold on and off
	public final static boolean BOARD_FONT_BOLD = true;

	// BOARD_BORDER_COLOR sets the border color of the board
	public final static Color BOARD_BORDER_COLOR = Color.black;

	// BOARD_BORDER_INPLAY_COLOR sets the color of the borders that are in
	// play(squares with letters)
	public final static Color BOARD_BORDER_INPLAY_COLOR = Color.black;

	// BOARD_BORDER_NOTINPLAY_COLOR sets the color of the borders that are not in
	// play.
	public final static Color BOARD_BORDER_NOTINPLAY_COLOR = Color.white;

	// BOARD_BACKGROUND_COLOR sets the background color of the board
	public final static Color BOARD_BACKGROUND_COLOR = Color.white;

	// SLIDE_NUMBER_COLOR sets the color of the slide number
	public final static Color SLIDE_NUMBER_COLOR = Color.blue;

	// SLIDE_NUMBER_FONTSIZE change the size of the number
	public final static double SLIDE_NUMBER_FONTSIZE = 45.0;

	// SLIDE_NUMBER_FONTBOLD sets the slide number to bold or not
	public final static boolean SLIDE_NUMBER_FONTBOLD = true;

	// TABLE_FONT_SIZE changes the font size of the letters in the board.
	// Note: changing the font size can also change how big the board gets.
	public final static double TABLE_FONT_SIZE = 21.0;

	// SLIDE_WIDTH specifies how big the width of the ppt slide is.
	public final static int SLIDE_WIDTH = 1500;

	// SLIDE_HEIGHT specifies what the height of the ppt slide is.
	public final static int SLIDE_HEIGHT = 1300;

	// BOARD_ROW specifies how many rows you want in your board.
	// Note: changing this will affect how it appears on the ppt.
	public final static int TABLE_ROW = 16;

	// BOARD_COL specifies how many columns you want in your board.
	// Note: changing this will affect how it appears on the ppt.
	public final static int TABLE_COL = 24;
	
	public final static double TABLE_COL_WIDTH = 60.0;
	
	public final static double TABLE_ROW_HEIGHT = 60.0;

	// TITLE_FONT_SIZE specifies the font size of the title.
	public final static double TITLE_FONT_SIZE = 40.0;

	// TITLE_FONT_COLOR changes the color of the title.
	public final static Color TITLE_FONT_COLOR = Color.red;

	// TITLE_NAME specifies the name of the title on the slides.
	public final static String TITLE_NAME = "Crossword Puzzle";

	// TITLE_SOLUTION_NAME specifies the title of the solution slides.
	public final static String TITLE_SOLUTION_NAME = "Crossword Solution";

	// TITLE_SOLUTION_NAME specifies the title of the solution slides.
	public final static String TITLE_CLUE_NAME = "Crossword Clues";

	// TITLE_BOLD specifies if you want the title to be bold or not
	public final static boolean TITLE_BOLD = true;

	// GRID_WIDTH specifies the width of the grid/board
	public final static int GRID_WIDTH = 1150;

	// GRID_HEIGHT specifies the height of the grid/board
	public final static int GRID_HEIGHT = 1150;

	// LOGO_WIDTH specifies the width of the logo on each slide.
	public final static int LOGO_WIDTH = 300;

	// LOGO_HEIGHT specifies the height of the logo on each slide.
	public final static int LOGO_HEIGHT = 100;
	
	public final static Color CLUE_FONT_COLOR = Color.blue;
	
	public final static double CLUE_FONT_SIZE = 20.0;
	
	public final static int CLUE_BOX_WIDTH = 1200;
	
	public final static int CLUE_BOX_HEIGHT = 900;

}
