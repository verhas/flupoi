package javax0.flupoi;

import java.util.StringTokenizer;

public class RangeBuilder {

	private Range range;
	private String rangeString;

	protected RangeBuilder(String rangeString) {
		this.rangeString = rangeString;
	}

	private String rangeStartString;
	private String rangeEndString;

	/**
	 * Nx:My or (x,y):(x,y)
	 */
	protected Range build() {
		range = new Range();
		removeAllSpacesFromRangeString();
		lowerCaseRangeString();
		separateRangeStartAndRangeEndStrings();
		range.setStart(convert(rangeStartString));
		if (rangeEndString != null) {
			range.setEnd(convert(rangeEndString));
		} else {
			range.setEnd(range.getStart());
		}
		range.setDirection(calculateDirection());
		return range;
	}

	private void removeAllSpacesFromRangeString() {
		rangeString = rangeString.replaceAll("\\s+", "");
	}

	private void lowerCaseRangeString() {
		rangeString = rangeString.toLowerCase();
	}

	private boolean isAlphaFormat(final String coordinate) {
		return coordinate.charAt(0) != '(';
	}

	private void separateRangeStartAndRangeEndStrings() {
		final StringTokenizer st = new StringTokenizer(rangeString, ":");
		rangeStartString = st.nextToken();
		if (st.hasMoreTokens()) {
			rangeEndString = st.nextToken();
		}
	}

	private Point convertNumericFormat(final String coordinate) {
		StringBuilder sb = new StringBuilder(coordinate);
		dropCharacter(sb); // the '('
		Point point = new Point();
		convertNumericOrJoker(sb, point.getX());
		if (sb.charAt(0) != ',') {
			throw new IllegalArgumentException(
					"'"
							+ coordinate
							+ "' is not a valid coordinate, there is no separating comma");
		}
		dropCharacter(sb);
		convertNumericOrJoker(sb, point.getY());
		if (sb.charAt(0) != ')') {
			throw new IllegalArgumentException(
					"'"
							+ coordinate
							+ "' is not a valid coordinate, does not finish with a closing parenthese");
		}
		dropCharacter(sb);
		if (sb.length() > 0) {
			throw new IllegalArgumentException("'" + coordinate
					+ "' is not a valid coordinate");
		}
		return point;
	}

	private int convertAlpha(StringBuilder sb) {
		int x = 0;
		while (sb.length() > 0 && Character.isAlphabetic(sb.charAt(0))) {
			x = 26 * x + (int) (sb.charAt(0) - 'a');
			dropCharacter(sb);
		}
		return x;
	}

	private int convertNumeric(StringBuilder sb) {
		if (sb.length() == 0) {
			throw new IllegalArgumentException();
		}
		int y = 0;
		while (sb.length() > 0 && Character.isDigit(sb.charAt(0))) {
			y = 10 * y + (int) (sb.charAt(0) - '0');
			dropCharacter(sb);
		}
		return y;
	}

	private char fetchSignumCharacter(StringBuilder sb) {
		char c = sb.charAt(0);
		if (c == '+' || c == '-') {
			dropCharacter(sb);
		}
		return c;
	}

	private void dropCharacter(StringBuilder sb) {
		sb.deleteCharAt(0);
	}

	private boolean firstCharacterIsJoker(StringBuilder sb) {
		return sb.length() > 0 && sb.charAt(0) == '*';
	}

	private void convertAlphaOrJoker(StringBuilder sb, Point point) {
		if (firstCharacterIsJoker(sb)) {
			dropCharacter(sb);
		} else {
			point.getX().setRelativity(fetchSignumCharacter(sb));
			final int x = convertAlpha(sb);
			point.getX().setValue(x);
		}
	}

	private void convertNumericOrJoker(StringBuilder sb, Coordinate coordinate) {
		if (firstCharacterIsJoker(sb)) {
			dropCharacter(sb);
		} else {
			coordinate.setRelativity(fetchSignumCharacter(sb));
			coordinate.setValue(convertNumeric(sb));
		}
	}

	/**
	 * XLS numbering goes from 1 but POI indexing goes from 0. For example A1 is
	 * (0,0)
	 * 
	 * @param point
	 *            to correct
	 */
	private void correctYCoordinateForIndexing(Point point) {
		if (point.getY().isDefined()) {
			point.getY().setValue(point.getY().getValue() - 1);
		}
	}

	private Point convertAlphaFormat(final String coordinate) {
		Point point = new Point();
		StringBuilder cord = new StringBuilder(coordinate);
		convertAlphaOrJoker(cord, point);
		convertNumericOrJoker(cord, point.getY());
		correctYCoordinateForIndexing(point);
		if (cord.length() > 0) {
			throw new IllegalArgumentException("'" + coordinate
					+ "' is not a valid coordinate.");
		}
		return point;
	}

	private Point convert(final String coordinate) {
		try {
			final Point point;
			if (isAlphaFormat(coordinate)) {
				point = convertAlphaFormat(coordinate);
			} else {
				point = convertNumericFormat(coordinate);
			}
			return point;
		} catch (IllegalArgumentException e) {
			if (e.getMessage() != null) {
				throw e;
			} else {
				throw new IllegalArgumentException(coordinate
						+ " is invalid coordinate");
			}
		}

	}

	private RangeDirection calculateDirection() {
		RangeDirection dir = null;
		int joker = 0;
		if (!range.getStart().getX().isDefined()) {
			dir = RangeDirection.LEFT;
			joker++;
		}
		if (!range.getStart().getY().isDefined()) {
			dir = RangeDirection.UP;
			joker++;
		}
		if (!range.getEnd().getX().isDefined()) {
			dir = RangeDirection.RIGHT;
			joker++;
		}
		if (!range.getEnd().getY().isDefined()) {
			dir = RangeDirection.DOWN;
			joker++;
		}
		if (joker > 1) {
			throw new IllegalArgumentException("'" + rangeString
					+ "' is invalid.");
		}
		return dir;
	}
}
