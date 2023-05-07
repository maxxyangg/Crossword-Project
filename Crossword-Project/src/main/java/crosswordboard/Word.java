package crosswordboard;


import java.util.ArrayList;

public class Word {
	private String word;
	private ArrayList<Letter> letters;
	private int row;
	private int col;
	private String direction;
	private String clue;
	private int clueNumber;

	public Word(String word, ArrayList<Letter> letters, int row, int col, String direction, String clue, int clueNumber) {
		this.word = word;
		this.letters = letters;
		this.row = row;
		this.col = col;
		this.direction = direction;
		this.clue = clue;
		this.clueNumber = clueNumber;
	}
	
	public Word() {
	}

	
	public String getClue() {
		return clue;
	}

	public void setClue(String clue) {
		this.clue = clue;
	}

	public String getWord() {
		return word;
	}

	public void setWord(String word) {
		this.word = word;
	}

	public ArrayList<Letter> getLetters() {
		return letters;
	}

	public void setLetters(ArrayList<Letter> letters) {
		this.letters = letters;
	}

	public int getRow() {
		return row;
	}

	public void setRow(int row) {
		this.row = row;
	}

	public int getCol() {
		return col;
	}

	public void setCol(int col) {
		this.col = col;
	}

	public String getDirection() {
		return direction;
	}

	public void setDirection(String direction) {
		this.direction = direction;
	}

	public int getClueNumber() {
		return clueNumber;
	}

	public void setClueNumber(int clueNumber) {
		this.clueNumber = clueNumber;
	}

	
	
}
